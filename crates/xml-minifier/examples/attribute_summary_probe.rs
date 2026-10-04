//! Small native timing guard for the XML attribute-summary candidate.
//!
//! The probe keeps its corpus and expected audit results outside the timed
//! batches. Each batch black-boxes the input and every returned result, then
//! checks its last result against the untimed expected result before emitting
//! one TSV row. The output's comment rows carry the exact corpus bytes and
//! stable `Debug` verdicts so baseline and candidate runs can be matched.
//!
//! ```text
//! cargo run -p xml-minifier --example attribute_summary_probe --release
//! ```

#![allow(
    clippy::print_stdout,
    reason = "the probe's TSV is its intended command-line output"
)]

use std::fmt::{self, Display, Formatter};
use std::hint::black_box;
use std::time::Instant;

use xml_minifier::audit::{self, Error, Limits, Report};

type AuditResult = Result<Report, Error>;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Policy {
    Source,
    Compact,
    Authored,
}

impl Policy {
    const ALL: [Self; 3] = [Self::Source, Self::Compact, Self::Authored];

    const fn index(self) -> usize {
        match self {
            Self::Source => 0,
            Self::Compact => 1,
            Self::Authored => 2,
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::Source => "source",
            Self::Compact => "compact",
            Self::Authored => "authored",
        }
    }
}

#[derive(Debug)]
struct ProbeError(String);

impl Display for ProbeError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
    }
}

impl std::error::Error for ProbeError {}

#[derive(Debug)]
struct Config {
    samples: usize,
    warmup: usize,
    iterations: usize,
    reverse: bool,
}

struct Case {
    name: &'static str,
    input: Vec<u8>,
    input_hex: String,
    limits: Limits,
    expected: [AuditResult; 3],
}

#[derive(Clone, Copy)]
struct Job {
    case: usize,
    policy: Policy,
}

struct Row {
    case: &'static str,
    policy: &'static str,
    sample: usize,
    elapsed_ns: u128,
    iterations: usize,
}

fn main() -> Result<(), ProbeError> {
    let config = parse_args()?;
    let cases = build_cases()?;
    let mut jobs = Vec::with_capacity(cases.len() * Policy::ALL.len());
    for case in 0..cases.len() {
        for policy in Policy::ALL {
            jobs.push(Job { case, policy });
        }
    }
    if config.reverse {
        jobs.reverse();
    }

    let mut rows = Vec::with_capacity(jobs.len() * config.samples);
    for job in jobs {
        let case = &cases[job.case];
        let expected = &case.expected[job.policy.index()];
        for _ in 0..config.warmup {
            let (_elapsed_ns, last_result) =
                timed_batch(job.policy, &case.input, case.limits, config.iterations)?;
            check_result(case.name, job.policy, expected, last_result)?;
        }
        for sample in 0..config.samples {
            let (elapsed_ns, last_result) =
                timed_batch(job.policy, &case.input, case.limits, config.iterations)?;
            check_result(case.name, job.policy, expected, last_result)?;
            rows.push(Row {
                case: case.name,
                policy: job.policy.name(),
                sample,
                elapsed_ns,
                iterations: config.iterations,
            });
        }
    }

    println!("# attribute_summary_probe_v1");
    println!(
        "# samples={} warmup={} reverse={}",
        config.samples, config.warmup, config.reverse
    );
    for case in &cases {
        println!(
            "# input\tcase={}\tbytes={}\thex={}",
            case.name,
            case.input.len(),
            case.input_hex
        );
        for policy in Policy::ALL {
            let expected = &case.expected[policy.index()];
            println!(
                "# outcome\tcase={}\tpolicy={}\texpected={}",
                case.name,
                policy.name(),
                debug_one_line(expected)
            );
        }
    }
    println!("case\tpolicy\tsample\telapsed_ns\titerations");
    for row in rows {
        println!(
            "{}\t{}\t{}\t{}\t{}",
            row.case, row.policy, row.sample, row.elapsed_ns, row.iterations
        );
    }
    Ok(())
}

fn parse_args() -> Result<Config, ProbeError> {
    let mut samples = 30usize;
    let mut warmup = 5usize;
    let mut iterations = 1_000usize;
    let mut reverse = false;
    let mut arguments = std::env::args().skip(1);
    while let Some(argument) = arguments.next() {
        if argument == "--reverse" {
            reverse = true;
            continue;
        }
        if argument == "--help" || argument == "-h" {
            println!(
                "usage: attribute_summary_probe [--samples N] [--warmup N] [--iterations N] [--reverse]"
            );
            std::process::exit(0);
        }
        let (name, inline) = match argument.split_once('=') {
            Some((name, value)) => (name, Some(value.to_owned())),
            None => (argument.as_str(), None),
        };
        let value = match inline {
            Some(value) => value,
            None => arguments
                .next()
                .ok_or_else(|| ProbeError(format!("missing value for {name}")))?,
        };
        let parsed = value
            .parse::<usize>()
            .map_err(|error| ProbeError(format!("invalid {name} value {value:?}: {error}")))?;
        match name {
            "--samples" if parsed > 0 => samples = parsed,
            "--warmup" => warmup = parsed,
            "--iterations" if parsed > 0 => iterations = parsed,
            "--samples" | "--iterations" => {
                return Err(ProbeError(format!("{name} must be greater than zero")));
            },
            _ => return Err(ProbeError(format!("unknown argument {name}"))),
        }
    }
    Ok(Config {
        samples,
        warmup,
        iterations,
        reverse,
    })
}

fn build_cases() -> Result<Vec<Case>, ProbeError> {
    let specifications = [
        ("cardinality-0", b"<root/>".to_vec()),
        ("cardinality-1", attributes(1)),
        ("cardinality-2", attributes(2)),
        ("cardinality-8", attributes(8)),
        ("cardinality-32", attributes(32)),
        ("cardinality-33", attributes(33)),
        ("cardinality-128", attributes(128)),
        (
            "xml-space-valid",
            b"<root xml:space=\"preserve\"/>".to_vec(),
        ),
        ("xml-space-invalid", b"<root xml:space=\"maybe\"/>".to_vec()),
        ("duplicate", b"<root a=\"1\" a=\"2\"/>".to_vec()),
        (
            "malformed-tail",
            b"<root a=\"1\" b=\"unterminated/>".to_vec(),
        ),
        ("malformed-second", b"<root a=\"1\" b/>".to_vec()),
    ];
    specifications
        .into_iter()
        .map(|(name, input)| {
            let limits = Limits::new(
                input.len(),
                64,
                4_096,
                1_024,
                input.len().max(1),
                input.len().max(1),
            )
            .map_err(|error| ProbeError(format!("limits for {name}: {error}")))?;
            let expected = [
                audit::verify_source(&input, limits),
                audit::verify(&input, limits),
                audit::verify_authored(&input, limits),
            ];
            Ok(Case {
                name,
                input_hex: hex(&input),
                input,
                limits,
                expected,
            })
        })
        .collect()
}

fn attributes(count: usize) -> Vec<u8> {
    let mut xml = String::from("<root");
    for index in 0..count {
        xml.push_str(" a");
        xml.push_str(&index.to_string());
        xml.push_str("=\"");
        xml.push_str(&index.to_string());
        xml.push_str("\"");
    }
    xml.push_str("/>");
    xml.into_bytes()
}

fn hex(bytes: &[u8]) -> String {
    const DIGITS: &[u8; 16] = b"0123456789abcdef";
    let mut output = String::with_capacity(bytes.len() * 2);
    for &byte in bytes {
        output.push(char::from(DIGITS[usize::from(byte >> 4)]));
        output.push(char::from(DIGITS[usize::from(byte & 0x0f)]));
    }
    output
}

fn audit(policy: Policy, input: &[u8], limits: Limits) -> AuditResult {
    match policy {
        Policy::Source => audit::verify_source(input, limits),
        Policy::Compact => audit::verify(input, limits),
        Policy::Authored => audit::verify_authored(input, limits),
    }
}

fn timed_batch(
    policy: Policy,
    input: &[u8],
    limits: Limits,
    iterations: usize,
) -> Result<(u128, AuditResult), ProbeError> {
    let started = Instant::now();
    let mut last_result = None;
    for _ in 0..iterations {
        let result = audit(policy, black_box(input), limits);
        last_result = Some(black_box(result));
    }
    let elapsed_ns = started.elapsed().as_nanos();
    // `iterations` is validated as nonzero by the CLI parser.
    let last_result =
        last_result.ok_or_else(|| ProbeError("timed batch had no iterations".to_owned()))?;
    Ok((elapsed_ns, last_result))
}

fn check_result(
    case: &str,
    policy: Policy,
    expected: &AuditResult,
    last_result: AuditResult,
) -> Result<(), ProbeError> {
    let last_result = black_box(last_result);
    if &last_result != expected {
        return Err(ProbeError(format!(
            "audit result changed for case {case:?}, policy {}: expected {}, got {}",
            policy.name(),
            debug_one_line(expected),
            debug_one_line(&last_result)
        )));
    }
    Ok(())
}

fn debug_one_line<T: fmt::Debug>(value: &T) -> String {
    format!("{value:?}")
        .replace('\t', "\\t")
        .replace('\r', "\\r")
        .replace('\n', "\\n")
}
