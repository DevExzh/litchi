//! Direct evidence probe for the checked-attribute boundary.
//!
//! The two helper implementations are local modules copied byte-for-byte from
//! the production baseline and the 0799 two-attribute candidate.  This binary is
//! intentionally independent of the workspace so both legs run in one
//! process, over the same prebuilt `BytesStart` inputs.

#[path = "baseline.rs"]
mod baseline;
#[path = "candidate.rs"]
mod candidate;

use std::env;
use std::error::Error;
use std::fs;
use std::hint::black_box;
use std::path::PathBuf;
use std::time::Instant;

use quick_xml::events::BytesStart;
use quick_xml::events::attributes::{AttrError, Attribute};
use serde::Serialize;

const SCHEMA: &str = "litchi.attribute-boundary-probe.v1";
const TOOL: &str = "attribute-boundary-probe-0799";
const DEFAULT_SAMPLES: usize = 1;
const DEFAULT_WARMUP: usize = 0;
const DEFAULT_ITERATIONS: usize = 1;
const MAX_SAMPLES: usize = 100_000;
const MAX_WARMUP: usize = 100_000;
const MAX_ITERATIONS: usize = 1_000_000;
const CLONE_TRANSITIONS: &[usize] = &[0, 1, 2, 3, 4, 5, 32, 33];
const LONG_VALUE_LENGTH: usize = 4096;
const FNV_OFFSET: u64 = 0xcbf2_9ce4_8422_2325;
const FNV_PRIME: u64 = 0x0000_0100_0000_01b3;
const CONSTRUCT_STEP: u64 = 0x9e37_79b9_7f4a_7c15;
const CONSTRUCT_SEED: u64 = 0x6a09_e667_f3bc_c909;
const VALUE_SEED: u64 = 0xbb67_ae85_84ca_a73b;

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "kebab-case")]
enum Leg {
    Before,
    After,
}

impl Leg {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            "before" => Ok(Self::Before),
            "after" => Ok(Self::After),
            _ => Err(format!("--leg must be before or after (got {value:?})").into()),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "kebab-case")]
enum Mode {
    Construct,
    Consume,
}

impl Mode {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            "construct" => Ok(Self::Construct),
            "consume" => Ok(Self::Consume),
            _ => Err(format!("--mode must be construct or consume (got {value:?})").into()),
        }
    }

    const fn timing_scope(self) -> &'static str {
        match self {
            Self::Construct => {
                "selected named construction owner call, including its common call dispatch"
            },
            Self::Consume => {
                "selected named consumption owner call, including its common call dispatch"
            },
        }
    }
}

#[derive(Clone, Debug)]
struct CaseSpec {
    id: String,
    category: String,
    attribute_count: Option<usize>,
    input: String,
}

#[derive(Clone, Debug, Serialize)]
struct SourceIdentity {
    bytes: usize,
    encoding: &'static str,
    value: String,
}

impl SourceIdentity {
    fn from_input(input: &str) -> Self {
        if input.len() <= 512 {
            Self {
                bytes: input.len(),
                encoding: "utf8",
                value: input.to_owned(),
            }
        } else {
            Self {
                bytes: input.len(),
                encoding: "hex",
                value: hex(input.as_bytes()),
            }
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize)]
#[serde(tag = "kind")]
enum ErrorRecord {
    ExpectedEq {
        position: usize,
    },
    ExpectedValue {
        position: usize,
    },
    UnquotedValue {
        position: usize,
    },
    ExpectedQuote {
        position: usize,
        quote: u8,
    },
    Duplicated {
        position: usize,
        first_position: usize,
    },
}

fn error_record(error: &AttrError) -> ErrorRecord {
    match error {
        AttrError::ExpectedEq(position) => ErrorRecord::ExpectedEq {
            position: *position,
        },
        AttrError::ExpectedValue(position) => ErrorRecord::ExpectedValue {
            position: *position,
        },
        AttrError::UnquotedValue(position) => ErrorRecord::UnquotedValue {
            position: *position,
        },
        AttrError::ExpectedQuote(position, quote) => ErrorRecord::ExpectedQuote {
            position: *position,
            quote: *quote,
        },
        AttrError::Duplicated(position, first_position) => ErrorRecord::Duplicated {
            position: *position,
            first_position: *first_position,
        },
    }
}

fn error_marker(error: &ErrorRecord) -> u64 {
    let (kind, first, second) = match error {
        ErrorRecord::ExpectedEq { position } => (1_u64, *position as u64, 0),
        ErrorRecord::ExpectedValue { position } => (2, *position as u64, 0),
        ErrorRecord::UnquotedValue { position } => (3, *position as u64, 0),
        ErrorRecord::ExpectedQuote { position, quote } => (4, *position as u64, *quote as u64),
        ErrorRecord::Duplicated {
            position,
            first_position,
        } => (5, *position as u64, *first_position as u64),
    };
    kind.wrapping_mul(CONSTRUCT_STEP) ^ first.rotate_left(17) ^ second.rotate_left(31)
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum OwnedItem {
    Attribute { key: Vec<u8>, value: Vec<u8> },
    Error(ErrorRecord),
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct OwnedTrace {
    items: Vec<OwnedItem>,
    repeated_none_calls: usize,
}

#[derive(Clone, Debug)]
struct FailFast<I> {
    inner: I,
    done: bool,
}

impl<'a, I> Iterator for FailFast<I>
where
    I: Iterator<Item = Result<Attribute<'a>, AttrError>>,
{
    type Item = Result<Attribute<'a>, AttrError>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        match self.inner.next() {
            Some(Err(error)) => {
                self.done = true;
                Some(Err(error))
            },
            Some(Ok(attribute)) => Some(Ok(attribute)),
            None => {
                self.done = true;
                None
            },
        }
    }
}

#[derive(Clone, Debug, Serialize)]
#[serde(tag = "kind")]
enum TraceItemSummary {
    Attribute {
        key_bytes: usize,
        key_hash: u64,
        value_bytes: usize,
        value_hash: u64,
    },
    Error {
        error: ErrorRecord,
    },
}

#[derive(Clone, Debug, Serialize)]
struct TraceSummary {
    accepted: usize,
    first_error: Option<ErrorRecord>,
    repeated_none_calls: usize,
    sequence_hash: u64,
    items: Vec<TraceItemSummary>,
}

impl OwnedTrace {
    fn summary(&self) -> TraceSummary {
        let first_error = self.items.iter().find_map(|item| match item {
            OwnedItem::Error(error) => Some(error.clone()),
            OwnedItem::Attribute { .. } => None,
        });
        let accepted = self
            .items
            .iter()
            .filter(|item| matches!(item, OwnedItem::Attribute { .. }))
            .count();
        let items = self
            .items
            .iter()
            .map(|item| match item {
                OwnedItem::Attribute { key, value } => TraceItemSummary::Attribute {
                    key_bytes: key.len(),
                    key_hash: hash_bytes(key),
                    value_bytes: value.len(),
                    value_hash: hash_bytes(value),
                },
                OwnedItem::Error(error) => TraceItemSummary::Error {
                    error: error.clone(),
                },
            })
            .collect();
        TraceSummary {
            accepted,
            first_error,
            repeated_none_calls: self.repeated_none_calls,
            sequence_hash: sequence_hash(&self.items),
            items,
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct CloneCheck {
    advance: usize,
    quick_xml_sequence_hash: u64,
    baseline_sequence_hash: u64,
    candidate_sequence_hash: u64,
    baseline_matches_quick_xml: bool,
    candidate_matches_quick_xml: bool,
    terminal_behavior_matches: bool,
}

#[derive(Clone, Debug, Serialize)]
struct OracleVerification {
    quick_xml: TraceSummary,
    baseline: TraceSummary,
    candidate: TraceSummary,
    baseline_matches_quick_xml: bool,
    candidate_matches_quick_xml: bool,
    clone_checks: Vec<CloneCheck>,
    all_checks_passed: bool,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize)]
struct LoopResult {
    checksum: u64,
    accepted: u64,
    error_marker: u64,
}

#[derive(Clone, Debug, Serialize)]
struct IteratorSizes {
    baseline_checked_attributes: usize,
    candidate_checked_attributes: usize,
}

#[derive(Clone, Debug, Serialize)]
struct SampleRecord {
    index: usize,
    elapsed_ns: u64,
    checksum: u64,
    accepted: u64,
    error_marker: u64,
}

#[derive(Clone, Debug, Serialize)]
struct Report {
    schema: &'static str,
    tool: &'static str,
    binary: String,
    leg: Leg,
    case: String,
    category: String,
    attribute_count: Option<usize>,
    mode: Mode,
    iterations: usize,
    warmup: usize,
    samples_requested: usize,
    timing_scope: &'static str,
    source: SourceIdentity,
    semantic_oracle: OracleVerification,
    iterator_sizes: IteratorSizes,
    expected_result: LoopResult,
    samples: Vec<SampleRecord>,
}

#[derive(Clone, Debug, Serialize)]
struct ListedCase {
    id: String,
    category: String,
    attribute_count: Option<usize>,
    source: SourceIdentity,
    expected_baseline: TraceSummary,
}

#[derive(Clone, Debug, Serialize)]
struct SelfCheckCase {
    id: String,
    all_checks_passed: bool,
    baseline_matches_quick_xml: bool,
    candidate_matches_quick_xml: bool,
    clone_checks: Vec<CloneCheck>,
}

#[derive(Clone, Debug, Serialize)]
struct SelfCheckReport {
    schema: &'static str,
    tool: &'static str,
    cases: Vec<SelfCheckCase>,
    all_checks_passed: bool,
}

#[derive(Debug)]
struct Options {
    leg: Option<Leg>,
    case_id: Option<String>,
    mode: Option<Mode>,
    samples: usize,
    warmup: usize,
    iterations: usize,
    output: Option<PathBuf>,
    list_cases: bool,
    self_check: bool,
    help: bool,
}

fn main() -> Result<(), Box<dyn Error>> {
    let options = parse_args(env::args().skip(1))?;
    if options.help {
        print_usage();
        return Ok(());
    }
    if options.list_cases {
        return print_list_cases();
    }
    if options.self_check {
        return run_self_check();
    }
    run_timed(options)
}

fn parse_args<I>(args: I) -> Result<Options, Box<dyn Error>>
where
    I: IntoIterator<Item = String>,
{
    let mut options = Options {
        leg: None,
        case_id: None,
        mode: None,
        samples: DEFAULT_SAMPLES,
        warmup: DEFAULT_WARMUP,
        iterations: DEFAULT_ITERATIONS,
        output: None,
        list_cases: false,
        self_check: false,
        help: false,
    };
    let mut args = args.into_iter();
    while let Some(argument) = args.next() {
        match argument.as_str() {
            "--help" | "-h" => options.help = true,
            "--list-cases" => options.list_cases = true,
            "--self-check" => options.self_check = true,
            "--leg" => options.leg = Some(Leg::parse(&required_arg(&mut args, "--leg")?)?),
            "--case" => options.case_id = Some(required_arg(&mut args, "--case")?),
            "--mode" => options.mode = Some(Mode::parse(&required_arg(&mut args, "--mode")?)?),
            "--samples" => {
                options.samples = parse_count(
                    &required_arg(&mut args, "--samples")?,
                    "--samples",
                    MAX_SAMPLES,
                )?;
            },
            "--warmup" => {
                options.warmup = parse_warmup(&required_arg(&mut args, "--warmup")?, MAX_WARMUP)?;
            },
            "--iterations" => {
                options.iterations = parse_count(
                    &required_arg(&mut args, "--iterations")?,
                    "--iterations",
                    MAX_ITERATIONS,
                )?;
            },
            "--output" => {
                options.output = Some(PathBuf::from(required_arg(&mut args, "--output")?))
            },
            other => return Err(format!("unknown argument {other:?}; use --help").into()),
        }
    }
    Ok(options)
}

fn required_arg<I>(args: &mut I, option: &str) -> Result<String, Box<dyn Error>>
where
    I: Iterator<Item = String>,
{
    args.next()
        .ok_or_else(|| format!("{option} requires a value").into())
}

fn parse_count(value: &str, option: &str, maximum: usize) -> Result<usize, Box<dyn Error>> {
    let parsed = value
        .parse::<usize>()
        .map_err(|error| format!("{option} must be a positive integer: {error}"))?;
    if parsed == 0 || parsed > maximum {
        return Err(format!("{option} must be in 1..={maximum} (got {parsed})").into());
    }
    Ok(parsed)
}

fn parse_warmup(value: &str, maximum: usize) -> Result<usize, Box<dyn Error>> {
    let parsed = value
        .parse::<usize>()
        .map_err(|error| format!("--warmup must be a non-negative integer: {error}"))?;
    if parsed > maximum {
        return Err(format!("--warmup must be <= {maximum} (got {parsed})").into());
    }
    Ok(parsed)
}

fn print_usage() {
    println!(
        "usage: attribute-boundary-probe --leg before|after --case CASE --mode construct|consume \
         --samples N --warmup N --iterations N --output FILE\n\
         attribute-boundary-probe --list-cases\n\
         attribute-boundary-probe --self-check"
    );
}

fn cases() -> Vec<CaseSpec> {
    let mut cases = Vec::new();
    for count in [0, 1, 2, 3, 4, 5, 8, 9, 16, 17, 32, 33, 64] {
        cases.push(CaseSpec {
            id: format!("distinct-{count}"),
            category: "distinct".to_owned(),
            attribute_count: Some(count),
            input: distinct_input(count),
        });
    }
    for count in [1, 2, 4, 5, 32, 33] {
        cases.push(CaseSpec {
            id: format!("duplicate-valid-after-{count}"),
            category: "duplicate-valid".to_owned(),
            attribute_count: Some(count + 1),
            input: duplicate_input(count, "n0=\"again\""),
        });
    }
    for (category, suffix) in [
        (
            "duplicate-long-quoted",
            format!("n0=\"{}\"", "x".repeat(LONG_VALUE_LENGTH)),
        ),
        (
            "duplicate-long-unterminated",
            format!("n0=\"{}", "x".repeat(LONG_VALUE_LENGTH)),
        ),
        (
            "duplicate-unquoted",
            format!("n0={}", "x".repeat(LONG_VALUE_LENGTH)),
        ),
    ] {
        for count in [1, 2, 33] {
            cases.push(CaseSpec {
                id: format!("{category}-after-{count}"),
                category: category.to_owned(),
                attribute_count: Some(count + 1),
                input: duplicate_input(count, &suffix),
            });
        }
    }
    for (category, suffix) in [
        ("syntax-flag", "flag"),
        ("syntax-unique-tail", "tail=x"),
        ("syntax-equals-value", "=\"1\""),
    ] {
        for count in [0, 2, 4, 33] {
            cases.push(CaseSpec {
                id: format!("{category}-after-{count}"),
                category: category.to_owned(),
                attribute_count: Some(count),
                input: syntax_input(count, suffix),
            });
        }
    }
    debug_assert_eq!(cases.len(), 39);
    cases
}

fn distinct_input(count: usize) -> String {
    let attributes = (0..count)
        .map(|index| format!("n{index}=\"{index}\""))
        .collect::<Vec<_>>();
    if attributes.is_empty() {
        "e".to_owned()
    } else {
        format!("e {}", attributes.join(" "))
    }
}

fn duplicate_input(count: usize, duplicate: &str) -> String {
    let prefix = distinct_input(count);
    format!("{prefix} {duplicate}")
}

fn syntax_input(count: usize, suffix: &str) -> String {
    let prefix = distinct_input(count);
    format!("{prefix} {suffix}")
}

fn make_tag(input: &str) -> BytesStart<'static> {
    let name_len = input.find([' ', '\t', '\r', '\n']).unwrap_or(input.len());
    BytesStart::from_content(input.to_owned(), name_len)
}

fn hash_bytes(bytes: &[u8]) -> u64 {
    let mut hash = FNV_OFFSET;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(FNV_PRIME);
    }
    hash
}

fn sequence_hash(items: &[OwnedItem]) -> u64 {
    let mut hash = FNV_OFFSET;
    for item in items {
        match item {
            OwnedItem::Attribute { key, value } => {
                hash = hash_byte(hash, 1);
                hash = hash_u64(hash, key.len() as u64);
                hash = hash_u64(hash, hash_bytes(key));
                hash = hash_u64(hash, value.len() as u64);
                hash = hash_u64(hash, hash_bytes(value));
            },
            OwnedItem::Error(error) => {
                hash = hash_byte(hash, 2);
                hash = hash_u64(hash, error_marker(error));
            },
        }
    }
    hash
}

fn hash_byte(mut hash: u64, byte: u8) -> u64 {
    hash ^= u64::from(byte);
    hash.wrapping_mul(FNV_PRIME)
}

fn hash_u64(mut hash: u64, value: u64) -> u64 {
    for byte in value.to_le_bytes() {
        hash = hash_byte(hash, byte);
    }
    hash
}

fn hex(bytes: &[u8]) -> String {
    const DIGITS: &[u8; 16] = b"0123456789abcdef";
    let mut output = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        output.push(DIGITS[usize::from(byte >> 4)] as char);
        output.push(DIGITS[usize::from(byte & 0x0f)] as char);
    }
    output
}

fn collect_trace<'a, I>(mut iterator: I) -> OwnedTrace
where
    I: Iterator<Item = Result<Attribute<'a>, AttrError>>,
{
    let mut items = Vec::new();
    loop {
        match iterator.next() {
            Some(Ok(attribute)) => items.push(OwnedItem::Attribute {
                key: attribute.key.as_ref().to_vec(),
                value: attribute.value.as_ref().to_vec(),
            }),
            Some(Err(error)) => {
                items.push(OwnedItem::Error(error_record(&error)));
                break;
            },
            None => break,
        }
    }
    let mut repeated_none_calls = 0;
    for _ in 0..2 {
        if iterator.next().is_none() {
            repeated_none_calls += 1;
        }
    }
    OwnedTrace {
        items,
        repeated_none_calls,
    }
}

#[allow(clippy::disallowed_methods)]
fn quick_xml_trace(tag: &BytesStart<'_>) -> OwnedTrace {
    collect_trace(FailFast {
        inner: tag.attributes(),
        done: false,
    })
}

fn baseline_trace(tag: &BytesStart<'_>) -> OwnedTrace {
    collect_trace(<BytesStart<'_> as baseline::BytesStartExt>::checked_attributes(tag))
}

fn candidate_trace(tag: &BytesStart<'_>) -> OwnedTrace {
    collect_trace(<BytesStart<'_> as candidate::BytesStartExt>::checked_attributes(tag))
}

fn advance<I>(iterator: &mut I, count: usize)
where
    I: Iterator,
{
    for _ in 0..count {
        let _ = iterator.next();
    }
}

#[allow(clippy::disallowed_methods)]
fn quick_xml_clone_trace(tag: &BytesStart<'_>, count: usize) -> OwnedTrace {
    let mut iterator = FailFast {
        inner: tag.attributes(),
        done: false,
    };
    advance(&mut iterator, count);
    collect_trace(iterator.clone())
}

fn baseline_clone_trace(tag: &BytesStart<'_>, count: usize) -> OwnedTrace {
    let mut iterator = <BytesStart<'_> as baseline::BytesStartExt>::checked_attributes(tag);
    advance(&mut iterator, count);
    collect_trace(iterator.clone())
}

fn candidate_clone_trace(tag: &BytesStart<'_>, count: usize) -> OwnedTrace {
    let mut iterator = <BytesStart<'_> as candidate::BytesStartExt>::checked_attributes(tag);
    advance(&mut iterator, count);
    collect_trace(iterator.clone())
}

fn oracle(tag: &BytesStart<'_>) -> OracleVerification {
    let quick_xml = quick_xml_trace(tag);
    let baseline = baseline_trace(tag);
    let candidate = candidate_trace(tag);
    let baseline_matches_quick_xml = baseline == quick_xml;
    let candidate_matches_quick_xml = candidate == quick_xml;
    let mut clone_checks = Vec::with_capacity(CLONE_TRANSITIONS.len());
    for advance_count in CLONE_TRANSITIONS {
        let quick_clone = quick_xml_clone_trace(tag, *advance_count);
        let baseline_clone = baseline_clone_trace(tag, *advance_count);
        let candidate_clone = candidate_clone_trace(tag, *advance_count);
        let terminal_behavior_matches = quick_clone.repeated_none_calls
            == baseline_clone.repeated_none_calls
            && quick_clone.repeated_none_calls == candidate_clone.repeated_none_calls;
        clone_checks.push(CloneCheck {
            advance: *advance_count,
            quick_xml_sequence_hash: sequence_hash(&quick_clone.items),
            baseline_sequence_hash: sequence_hash(&baseline_clone.items),
            candidate_sequence_hash: sequence_hash(&candidate_clone.items),
            baseline_matches_quick_xml: baseline_clone == quick_clone,
            candidate_matches_quick_xml: candidate_clone == quick_clone,
            terminal_behavior_matches,
        });
    }
    let all_checks_passed = baseline_matches_quick_xml
        && candidate_matches_quick_xml
        && clone_checks.iter().all(|check| {
            check.baseline_matches_quick_xml
                && check.candidate_matches_quick_xml
                && check.terminal_behavior_matches
        });
    OracleVerification {
        quick_xml: quick_xml.summary(),
        baseline: baseline.summary(),
        candidate: candidate.summary(),
        baseline_matches_quick_xml,
        candidate_matches_quick_xml,
        clone_checks,
        all_checks_passed,
    }
}

fn print_list_cases() -> Result<(), Box<dyn Error>> {
    let mut listed = Vec::with_capacity(39);
    for case in cases() {
        let tag = make_tag(&case.input);
        let verification = oracle(&tag);
        listed.push(ListedCase {
            id: case.id,
            category: case.category,
            attribute_count: case.attribute_count,
            source: SourceIdentity::from_input(&case.input),
            expected_baseline: verification.baseline,
        });
    }
    println!("{}", serde_json::to_string_pretty(&listed)?);
    Ok(())
}

fn run_self_check() -> Result<(), Box<dyn Error>> {
    let mut results = Vec::with_capacity(39);
    for case in cases() {
        let tag = make_tag(&case.input);
        let verification = oracle(&tag);
        results.push(SelfCheckCase {
            id: case.id,
            all_checks_passed: verification.all_checks_passed,
            baseline_matches_quick_xml: verification.baseline_matches_quick_xml,
            candidate_matches_quick_xml: verification.candidate_matches_quick_xml,
            clone_checks: verification.clone_checks,
        });
    }
    let all_checks_passed = results.iter().all(|case| case.all_checks_passed);
    let report = SelfCheckReport {
        schema: SCHEMA,
        tool: TOOL,
        cases: results,
        all_checks_passed,
    };
    println!("{}", serde_json::to_string_pretty(&report)?);
    if !all_checks_passed {
        return Err("attribute semantic self-check failed".into());
    }
    Ok(())
}

fn raw_error_marker(error: &AttrError) -> u64 {
    let (kind, first, second) = match error {
        AttrError::ExpectedEq(position) => (1_u64, *position as u64, 0),
        AttrError::ExpectedValue(position) => (2, *position as u64, 0),
        AttrError::UnquotedValue(position) => (3, *position as u64, 0),
        AttrError::ExpectedQuote(position, quote) => (4, *position as u64, *quote as u64),
        AttrError::Duplicated(position, first_position) => {
            (5, *position as u64, *first_position as u64)
        },
    };
    kind.wrapping_mul(CONSTRUCT_STEP) ^ first.rotate_left(17) ^ second.rotate_left(31)
}

fn attribute_checksum(key: &[u8], value: &[u8]) -> u64 {
    VALUE_SEED.rotate_left(5)
        ^ (key.len() as u64).wrapping_mul(CONSTRUCT_STEP)
        ^ (value.len() as u64).wrapping_mul(0xd6e8_feb8_6659_fd93)
}

fn drain_iterator<'a, I>(mut iterator: I) -> LoopResult
where
    I: Iterator<Item = Result<Attribute<'a>, AttrError>>,
{
    let mut checksum = VALUE_SEED;
    let mut accepted = 0_u64;
    let mut error_marker = 0_u64;
    loop {
        match iterator.next() {
            Some(Ok(attribute)) => {
                black_box(&attribute);
                accepted = accepted.wrapping_add(1);
                checksum = checksum.rotate_left(7)
                    ^ attribute_checksum(attribute.key.as_ref(), attribute.value.as_ref());
            },
            Some(Err(error)) => {
                // Keep the timed path allocation-free: the structured error
                // record is created only by the oracle outside the clock.
                black_box(&error);
                error_marker = raw_error_marker(&error);
                break;
            },
            None => break,
        }
    }
    LoopResult {
        checksum,
        accepted,
        error_marker,
    }
}

fn combine_iteration(total: &mut LoopResult, one: LoopResult, index: usize) {
    total.checksum = total
        .checksum
        .wrapping_add(one.checksum ^ (index as u64).wrapping_add(1).wrapping_mul(CONSTRUCT_STEP));
    total.accepted = total.accepted.wrapping_add(one.accepted);
    total.error_marker = total.error_marker.wrapping_add(one.error_marker);
}

#[inline(never)]
fn before_construct(tag: &BytesStart<'static>, iterations: usize) -> LoopResult {
    let mut result = LoopResult {
        checksum: CONSTRUCT_SEED,
        accepted: 0,
        error_marker: 0,
    };
    for index in 0..iterations {
        let iterator = <BytesStart<'_> as baseline::BytesStartExt>::checked_attributes(tag);
        black_box(&iterator);
        drop(iterator);
        result.checksum = result
            .checksum
            .wrapping_add((index as u64).wrapping_add(1).wrapping_mul(CONSTRUCT_STEP));
    }
    black_box(result)
}

#[inline(never)]
fn after_construct(tag: &BytesStart<'static>, iterations: usize) -> LoopResult {
    let mut result = LoopResult {
        checksum: CONSTRUCT_SEED,
        accepted: 0,
        error_marker: 0,
    };
    for index in 0..iterations {
        let iterator = <BytesStart<'_> as candidate::BytesStartExt>::checked_attributes(tag);
        black_box(&iterator);
        drop(iterator);
        result.checksum = result
            .checksum
            .wrapping_add((index as u64).wrapping_add(1).wrapping_mul(CONSTRUCT_STEP));
    }
    black_box(result)
}

#[inline(never)]
fn before_consume(tag: &BytesStart<'static>, iterations: usize) -> LoopResult {
    let mut result = LoopResult {
        checksum: CONSTRUCT_SEED,
        accepted: 0,
        error_marker: 0,
    };
    for index in 0..iterations {
        let iterator = <BytesStart<'_> as baseline::BytesStartExt>::checked_attributes(tag);
        let one = drain_iterator(iterator);
        combine_iteration(&mut result, one, index);
    }
    black_box(result)
}

#[inline(never)]
fn after_consume(tag: &BytesStart<'static>, iterations: usize) -> LoopResult {
    let mut result = LoopResult {
        checksum: CONSTRUCT_SEED,
        accepted: 0,
        error_marker: 0,
    };
    for index in 0..iterations {
        let iterator = <BytesStart<'_> as candidate::BytesStartExt>::checked_attributes(tag);
        let one = drain_iterator(iterator);
        combine_iteration(&mut result, one, index);
    }
    black_box(result)
}

fn expected_result(summary: &TraceSummary, mode: Mode, iterations: usize) -> LoopResult {
    match mode {
        Mode::Construct => {
            let mut checksum = CONSTRUCT_SEED;
            for index in 0..iterations {
                checksum = checksum
                    .wrapping_add((index as u64).wrapping_add(1).wrapping_mul(CONSTRUCT_STEP));
            }
            LoopResult {
                checksum,
                accepted: 0,
                error_marker: 0,
            }
        },
        Mode::Consume => {
            let mut one = LoopResult {
                checksum: VALUE_SEED,
                accepted: 0,
                error_marker: 0,
            };
            for item in &summary.items {
                match item {
                    TraceItemSummary::Attribute {
                        key_bytes,
                        value_bytes,
                        ..
                    } => {
                        one.accepted = one.accepted.wrapping_add(1);
                        one.checksum = one.checksum.rotate_left(7)
                            ^ attribute_checksum_from_lengths(*key_bytes, *value_bytes);
                    },
                    TraceItemSummary::Error { error } => {
                        one.error_marker = error_marker(error);
                    },
                }
            }
            let mut result = LoopResult {
                checksum: CONSTRUCT_SEED,
                accepted: 0,
                error_marker: 0,
            };
            for index in 0..iterations {
                combine_iteration(&mut result, one, index);
            }
            result
        },
    }
}

fn attribute_checksum_from_lengths(key_length: usize, value_length: usize) -> u64 {
    VALUE_SEED.rotate_left(5)
        ^ (key_length as u64).wrapping_mul(CONSTRUCT_STEP)
        ^ (value_length as u64).wrapping_mul(0xd6e8_feb8_6659_fd93)
}

type Runner = fn(&BytesStart<'static>, usize) -> LoopResult;

fn select_runner(leg: Leg, mode: Mode) -> Runner {
    match (leg, mode) {
        (Leg::Before, Mode::Construct) => before_construct,
        (Leg::After, Mode::Construct) => after_construct,
        (Leg::Before, Mode::Consume) => before_consume,
        (Leg::After, Mode::Consume) => after_consume,
    }
}

fn run_timed(options: Options) -> Result<(), Box<dyn Error>> {
    let leg = options
        .leg
        .ok_or("--leg is required unless --list-cases or --self-check is used")?;
    let mode = options
        .mode
        .ok_or("--mode is required unless --list-cases or --self-check is used")?;
    let case_id = options
        .case_id
        .ok_or("--case is required unless --list-cases or --self-check is used")?;
    let output = options
        .output
        .ok_or("--output is required unless --list-cases or --self-check is used")?;
    if output.exists() {
        return Err(format!("output already exists: {}", output.display()).into());
    }
    if let Some(parent) = output.parent()
        && !parent.as_os_str().is_empty()
        && !parent.is_dir()
    {
        return Err(format!("output parent is not a directory: {}", parent.display()).into());
    }

    let case = cases()
        .into_iter()
        .find(|case| case.id == case_id)
        .ok_or_else(|| format!("unknown case {case_id:?}; use --list-cases"))?;
    let tag = make_tag(&case.input);
    let semantic_oracle = oracle(&tag);
    if !semantic_oracle.all_checks_passed {
        return Err(format!("semantic oracle failed for case {}", case.id).into());
    }
    let expected = expected_result(&semantic_oracle.baseline, mode, options.iterations);
    let runner = select_runner(leg, mode);
    let mut samples = Vec::with_capacity(options.samples);

    for _ in 0..options.warmup {
        let result = runner(&tag, options.iterations);
        black_box(result);
        if result != expected {
            return Err(format!(
                "warmup checksum mismatch for {}: got {result:?}, expected {expected:?}",
                case.id
            )
            .into());
        }
    }
    for index in 0..options.samples {
        let started = Instant::now();
        let result = runner(&tag, options.iterations);
        let elapsed_ns = u64::try_from(started.elapsed().as_nanos())?;
        black_box(result);
        if result != expected {
            return Err(format!(
                "sample checksum mismatch for {}: got {result:?}, expected {expected:?}",
                case.id
            )
            .into());
        }
        samples.push(SampleRecord {
            index,
            elapsed_ns,
            checksum: result.checksum,
            accepted: result.accepted,
            error_marker: result.error_marker,
        });
    }

    let report = Report {
        schema: SCHEMA,
        tool: TOOL,
        binary: executable_name(),
        leg,
        case: case.id,
        category: case.category,
        attribute_count: case.attribute_count,
        mode,
        iterations: options.iterations,
        warmup: options.warmup,
        samples_requested: options.samples,
        timing_scope: mode.timing_scope(),
        source: SourceIdentity::from_input(&case.input),
        semantic_oracle,
        iterator_sizes: IteratorSizes {
            baseline_checked_attributes: std::mem::size_of::<baseline::CheckedAttributes<'static>>(
            ),
            candidate_checked_attributes: std::mem::size_of::<candidate::CheckedAttributes<'static>>(
            ),
        },
        expected_result: expected,
        samples,
    };
    fs::write(output, serde_json::to_vec_pretty(&report)?)?;
    Ok(())
}

fn executable_name() -> String {
    env::current_exe()
        .ok()
        .and_then(|path| {
            path.file_name()
                .map(|name| name.to_string_lossy().into_owned())
        })
        .unwrap_or_else(|| "attribute-boundary-probe".to_owned())
}
