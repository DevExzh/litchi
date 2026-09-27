//! Minimal fresh-writer lifecycle probe for change 0774.
//!
//! Each iteration constructs a new XLS writer, registers exactly 20,000 cells,
//! and publishes it with `write_to`. Formula iterations use the formula writer
//! API; numeric iterations provide the same cell-count control without formula
//! tokenization. The timer covers only construction, cell registration, and
//! `write_to`; output length and SHA-256 validation happen after the timer.
//!
//! Usage:
//!
//! ```text
//! xls-formula-probe --case formula|numeric --samples N --warmup N
//! ```

use std::io::Cursor;
use std::time::Instant;

use litchi_xls::writer::Writer;
use sha2::{Digest as _, Sha256};

const CELL_COUNT: usize = 20_000;

#[derive(Clone, Copy)]
enum Case {
    Formula,
    Numeric,
}

impl Case {
    fn parse(value: &str) -> Result<Self, String> {
        match value {
            "formula" => Ok(Self::Formula),
            "numeric" => Ok(Self::Numeric),
            other => Err(format!(
                "unknown case {other:?}; expected formula or numeric"
            )),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::Formula => "formula",
            Self::Numeric => "numeric",
        }
    }
}

struct Options {
    case: Case,
    samples: usize,
    warmup: usize,
}

fn usage() -> &'static str {
    "usage: xls-formula-probe --case formula|numeric --samples N --warmup N"
}

fn parse_count(flag: &str, value: Option<String>) -> Result<usize, String> {
    let value = value.ok_or_else(|| format!("missing value for {flag}; {usage()}"))?;
    let count = value
        .parse::<usize>()
        .map_err(|error| format!("invalid {flag} value {value:?}: {error}"))?;
    if flag == "--samples" && count == 0 {
        return Err("--samples must be at least 1".to_string());
    }
    Ok(count)
}

fn parse_options() -> Result<Options, String> {
    let mut case = None;
    let mut samples = None;
    let mut warmup = None;
    let mut arguments = std::env::args().skip(1);

    while let Some(flag) = arguments.next() {
        match flag.as_str() {
            "--case" => {
                let value = arguments
                    .next()
                    .ok_or_else(|| format!("missing value for --case; {}", usage()))?;
                case = Some(Case::parse(&value)?);
            },
            "--samples" => samples = Some(parse_count("--samples", arguments.next())?),
            "--warmup" => warmup = Some(parse_count("--warmup", arguments.next())?),
            "--help" | "-h" => return Err(usage().to_string()),
            other => return Err(format!("unknown argument {other:?}; {}", usage())),
        }
    }

    Ok(Options {
        case: case.ok_or_else(|| format!("missing --case; {}", usage()))?,
        samples: samples.ok_or_else(|| format!("missing --samples; {}", usage()))?,
        warmup: warmup.ok_or_else(|| format!("missing --warmup; {}", usage()))?,
    })
}

fn build_writer(case: Case) -> Result<Writer, Box<dyn std::error::Error>> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Probe")?;
    for row in 0..CELL_COUNT {
        let row = u32::try_from(row)?;
        match case {
            Case::Formula => {
                // Vary the row reference so each registration exercises the
                // complete formula-tokenization path while keeping the input
                // deterministic and within the BIFF8 grid.
                let formula = format!("A{}+1", row + 1);
                writer.write_formula(sheet, row, 0, &formula)?;
            },
            Case::Numeric => {
                writer.write_number(sheet, row, 0, f64::from(row) + 1.0)?;
            },
        }
    }
    Ok(writer)
}

fn one_sample(case: Case) -> Result<(u128, Vec<u8>), Box<dyn std::error::Error>> {
    let started = Instant::now();
    let mut writer = build_writer(case)?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    let elapsed_ns = started.elapsed().as_nanos();

    // Moving the completed buffer and validating it are deliberately after
    // the lifecycle timer. The bytes remain observable so the optimizer
    // cannot discard the publication.
    Ok((elapsed_ns, output.into_inner()))
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut text = String::with_capacity(digest.len() * 2);
    for byte in digest {
        use std::fmt::Write as _;
        write!(&mut text, "{byte:02x}").expect("writing to a String cannot fail");
    }
    text
}

fn observe(
    expected: &mut Option<(usize, String)>,
    bytes: Vec<u8>,
) -> Result<(), Box<dyn std::error::Error>> {
    let observed = (bytes.len(), sha256_hex(&bytes));
    match expected {
        None => *expected = Some(observed),
        Some(expected) if *expected == observed => {},
        Some(expected) => {
            return Err(format!(
                "non-deterministic output: expected {} bytes / {}, got {} bytes / {}",
                expected.0, expected.1, observed.0, observed.1
            )
            .into());
        },
    }
    std::hint::black_box(bytes);
    Ok(())
}

fn json_array(values: &[u128]) -> String {
    let mut text = String::from("[");
    for (index, value) in values.iter().enumerate() {
        if index != 0 {
            text.push(',');
        }
        text.push_str(&value.to_string());
    }
    text.push(']');
    text
}

fn run(options: Options) -> Result<(), Box<dyn std::error::Error>> {
    let mut expected = None;
    for _ in 0..options.warmup {
        let (_elapsed_ns, bytes) = one_sample(options.case)?;
        observe(&mut expected, bytes)?;
    }

    let mut samples_ns = Vec::with_capacity(options.samples);
    for _ in 0..options.samples {
        let (elapsed_ns, bytes) = one_sample(options.case)?;
        samples_ns.push(elapsed_ns);
        observe(&mut expected, bytes)?;
    }

    let (output_bytes, output_sha256) = expected.ok_or_else(|| {
        std::io::Error::new(std::io::ErrorKind::InvalidInput, "no samples were recorded")
    })?;
    println!(
        "{{\"schema\":\"xls-formula-probe-v1\",\"case\":\"{}\",\"cells\":{},\"warmup\":{},\"samples\":{},\"samples_ns\":{},\"output_bytes\":{},\"output_sha256\":\"{}\"}}",
        options.case.name(),
        CELL_COUNT,
        options.warmup,
        options.samples,
        json_array(&samples_ns),
        output_bytes,
        output_sha256,
    );
    Ok(())
}

fn main() {
    let result =
        parse_options().and_then(|options| run(options).map_err(|error| error.to_string()));
    if let Err(error) = result {
        eprintln!("{error}");
        std::process::exit(2);
    }
}
