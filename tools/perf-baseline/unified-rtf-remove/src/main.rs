//! Paired native and unified RTF paragraph-removal benchmark.
//!
//! Corpus construction, parsing, selector preparation, warm-up and output
//! validation are outside the timed intervals.  Each route reuses one warmed
//! immutable owner for the complete ABBA sample sequence.

use litchi_core::Position;
use litchi_core::selector::{Path, Query, Scope};
use litchi_ops::{File, OpenOptions, SourceMode, open_bytes_with};
use litchi_rtf::{EditOptions, Package, SemanticLimits as NativeLimits};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use std::error::Error;
use std::fmt::Write as _;
use std::hint::black_box;
use std::time::Instant;

type Result<T> = std::result::Result<T, Box<dyn Error + Send + Sync>>;

const DEFAULT_COUNTS: [usize; 2] = [10, 1_000];
const DEFAULT_WARMUP: usize = 2;
const DEFAULT_ROUNDS: usize = 9;
const NATIVE_MAX_OPERATIONS: usize = 10_000;
const NATIVE_MAX_REPLACEMENT_BYTES: usize = 16 * 1024 * 1024;

#[derive(Clone, Copy)]
enum Shape {
    SinglePath,
    QueryAll,
}

impl Shape {
    const fn name(self) -> &'static str {
        match self {
            Self::SinglePath => "single_path",
            Self::QueryAll => "query_all",
        }
    }
}

struct Config {
    counts: Vec<usize>,
    warmup: usize,
    rounds: usize,
}

fn corpus(count: usize) -> Result<Vec<u8>> {
    if count == 0 {
        return Err("the RTF corpus must contain at least one paragraph".into());
    }
    let mut source = String::with_capacity(count.saturating_mul(28).saturating_add(16));
    source.push_str(r"{\rtf1\ansi ");
    for position in 0..count {
        if position != 0 {
            source.push_str(r"\par ");
        }
        write!(source, "paragraph {position:05}")?;
    }
    source.push('}');
    Ok(source.into_bytes())
}

fn expected_texts(count: usize) -> Vec<String> {
    (0..count)
        .map(|position| format!("paragraph {position:05}"))
        .collect()
}

fn native_options() -> Result<EditOptions> {
    let limits = NativeLimits::builder()
        .max_operations(NATIVE_MAX_OPERATIONS)?
        .max_replacement_bytes(NATIVE_MAX_REPLACEMENT_BYTES)?
        .build()?;
    Ok(EditOptions::new().semantic_limits(limits))
}

fn open_unified(bytes: &[u8]) -> Result<File> {
    Ok(open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(SourceMode::Auto),
    )?)
}

fn scoped_paragraph_query() -> Result<Query<'static>> {
    let scope = Path::parse("/body")?.into_owned();
    Ok(Query::parse("p")?
        .into_owned()
        .with_scope(Scope::Path(Box::new(scope))))
}

fn native_once(
    owner: &Package,
    shape: Shape,
    position: usize,
    positions: &[Position],
    options: EditOptions,
) -> Result<Package> {
    let mut edit = owner.edit_with_options(options);
    match shape {
        Shape::SinglePath => {
            edit.remove_paragraph(Position::new(position))?;
        },
        Shape::QueryAll => {
            edit.remove_body_paragraphs(positions)?;
        },
    }
    Ok(edit.commit()?.into_package())
}

fn unified_once(
    owner: &File,
    shape: Shape,
    path: &Path<'static>,
    query: Option<Query<'static>>,
) -> Result<File> {
    match shape {
        Shape::SinglePath => Ok(owner.remove(path)?),
        Shape::QueryAll => Ok(owner.remove(query.ok_or("missing prepared query")?)?),
    }
}

fn native_texts(owner: &Package) -> Vec<String> {
    owner
        .body()
        .paragraphs()
        .map(|paragraph| paragraph.to_text())
        .collect()
}

fn unified_texts(owner: &File) -> Result<Vec<String>> {
    owner
        .query("p")?
        .map(|node| {
            let node = node?;
            Ok(node.view().text()?.unwrap_or_default().into_owned())
        })
        .collect::<Result<Vec<_>>>()
}

fn sha256(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

fn median(values: &[u128]) -> u128 {
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    sorted[sorted.len() / 2]
}

fn p75(values: &[u128]) -> u128 {
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    sorted[(sorted.len() * 3) / 4]
}

fn timing_summary(samples: &[u128]) -> Value {
    json!({
        "samples_ns": samples,
        "median_ns": median(samples),
        "p75_ns": p75(samples),
        "sample_count": samples.len(),
    })
}

fn output_oracle(
    native: &Package,
    unified: &File,
    source: &[u8],
    expected: &[String],
) -> Result<Value> {
    let native_bytes = native.to_bytes()?;
    let unified_bytes = unified.to_bytes()?;
    let native_text = native_texts(native);
    let unified_text = unified_texts(unified)?;
    if native_text != expected {
        return Err(format!(
            "native RTF removal text mismatch: observed {} paragraphs, expected {}",
            native_text.len(),
            expected.len()
        )
        .into());
    }
    if unified_text != expected {
        return Err(format!(
            "unified RTF removal text mismatch: observed {} paragraphs, expected {}",
            unified_text.len(),
            expected.len()
        )
        .into());
    }
    if native_bytes != unified_bytes {
        return Err(format!(
            "native and unified RTF removal outputs differ: native {} bytes/{}, unified {} bytes/{}",
            native_bytes.len(),
            sha256(&native_bytes),
            unified_bytes.len(),
            sha256(&unified_bytes),
        )
        .into());
    }
    if native_bytes == source {
        return Err("RTF removal produced unchanged bytes".into());
    }
    Ok(json!({
        "native_bytes": native_bytes.len(),
        "unified_bytes": unified_bytes.len(),
        "native_sha256": sha256(&native_bytes),
        "unified_sha256": sha256(&unified_bytes),
        "kind": "exact_bytes",
    }))
}

fn verify_native_sample(
    output: &Package,
    expected_bytes: &[u8],
    expected: &[String],
) -> Result<()> {
    let bytes = output.to_bytes()?;
    if bytes != expected_bytes {
        return Err("native timed RTF removal output changed from the warm oracle".into());
    }
    if native_texts(output) != expected {
        return Err("native timed RTF removal semantic output changed".into());
    }
    let reopened = Package::from_bytes(&bytes)?;
    if native_texts(&reopened) != expected {
        return Err("reopened native timed RTF output failed the semantic oracle".into());
    }
    Ok(())
}

fn verify_unified_sample(output: &File, expected_bytes: &[u8], expected: &[String]) -> Result<()> {
    let bytes = output.to_bytes()?;
    if bytes != expected_bytes {
        return Err("unified timed RTF removal output changed from the warm oracle".into());
    }
    if unified_texts(output)? != expected {
        return Err("unified timed RTF removal semantic output changed".into());
    }
    let reopened = open_unified(&bytes)?;
    if unified_texts(&reopened)? != expected {
        return Err("reopened unified timed RTF output failed the semantic oracle".into());
    }
    Ok(())
}

fn measure_case(
    source: &[u8],
    count: usize,
    shape: Shape,
    warmup: usize,
    rounds: usize,
) -> Result<Value> {
    let options = native_options()?;
    let position = count / 2;
    let positions = (0..count).map(Position::new).collect::<Vec<_>>();
    let path_text = format!("/body/p[{}]", position + 1);
    let path = Path::parse(&path_text)?.into_owned();
    let query = match shape {
        Shape::SinglePath => None,
        Shape::QueryAll => Some(scoped_paragraph_query()?),
    };
    let expected = match shape {
        Shape::SinglePath => {
            let mut values = expected_texts(count);
            values.remove(position);
            values
        },
        Shape::QueryAll => Vec::new(),
    };

    // Parsing and all selector construction stay outside the clocks.  The
    // two owners remain immutable and are reused for every timed operation.
    let native_owner = Package::from_bytes(source)?;
    let unified_owner = open_unified(source)?;
    let native_warm = native_once(&native_owner, shape, position, &positions, options)?;
    let unified_warm = match shape {
        Shape::SinglePath => unified_once(&unified_owner, shape, &path, None)?,
        Shape::QueryAll => unified_once(&unified_owner, shape, &path, query.clone())?,
    };
    let equivalence = output_oracle(&native_warm, &unified_warm, source, &expected)?;
    let expected_output = native_warm.to_bytes()?;
    verify_native_sample(&native_warm, &expected_output, &expected)?;
    verify_unified_sample(&unified_warm, &expected_output, &expected)?;
    drop((native_warm, unified_warm));

    // `File::remove` consumes an owned parsed query.  Clone all query values
    // before the timed region so neither query parsing nor target cloning is
    // charged to the unified route.
    let mut query_pool = query.map(|query| {
        (0..warmup
            .saturating_add(1)
            .saturating_add(rounds.saturating_mul(2)))
            .map(|_| query.clone())
            .collect::<Vec<_>>()
    });

    let warm_count = warmup;
    for _ in 0..warm_count {
        let native = native_once(
            black_box(&native_owner),
            shape,
            position,
            black_box(&positions),
            options,
        )?;
        black_box(&native);
        verify_native_sample(&native, &expected_output, &expected)?;
        drop(native);
        let unified = match shape {
            Shape::SinglePath => {
                unified_once(black_box(&unified_owner), shape, black_box(&path), None)?
            },
            Shape::QueryAll => unified_once(
                black_box(&unified_owner),
                shape,
                black_box(&path),
                Some(
                    query_pool
                        .as_mut()
                        .and_then(Vec::pop)
                        .ok_or("prepared query pool exhausted during warm-up")?,
                ),
            )?,
        };
        black_box(&unified);
        verify_unified_sample(&unified, &expected_output, &expected)?;
        drop(unified);
    }

    let mut native_ns = Vec::with_capacity(rounds * 2);
    let mut unified_ns = Vec::with_capacity(rounds * 2);
    for _ in 0..rounds {
        // ABBA keeps the two routes interleaved while the owners remain the
        // same warmed immutable snapshots.
        for native_leg in [true, false, false, true] {
            if native_leg {
                let start = Instant::now();
                let output = native_once(
                    black_box(&native_owner),
                    shape,
                    black_box(position),
                    black_box(&positions),
                    options,
                )?;
                native_ns.push(start.elapsed().as_nanos());
                black_box(&output);
                verify_native_sample(&output, &expected_output, &expected)?;
                drop(output);
            } else {
                let start = Instant::now();
                let output = match shape {
                    Shape::SinglePath => {
                        unified_once(black_box(&unified_owner), shape, black_box(&path), None)?
                    },
                    Shape::QueryAll => unified_once(
                        black_box(&unified_owner),
                        shape,
                        black_box(&path),
                        Some(
                            query_pool
                                .as_mut()
                                .and_then(Vec::pop)
                                .ok_or("prepared query pool exhausted during timing")?,
                        ),
                    )?,
                };
                unified_ns.push(start.elapsed().as_nanos());
                black_box(&output);
                verify_unified_sample(&output, &expected_output, &expected)?;
                drop(output);
            }
        }
    }

    let source_sha256 = sha256(source);
    let removed = match shape {
        Shape::SinglePath => 1,
        Shape::QueryAll => positions.len(),
    };
    Ok(json!({
        "shape": shape.name(),
        "paragraphs": count,
        "removed": removed,
        "selector": match shape {
            Shape::SinglePath => path_text.as_str(),
            Shape::QueryAll => "/body/p",
        },
        "corpus": {
            "bytes": source.len(),
            "sha256": source_sha256,
        },
        "warm_equivalence": equivalence,
        "timing": {
            "native": timing_summary(&native_ns),
            "unified": timing_summary(&unified_ns),
        },
        "protocol": {
            "warmup": warmup,
            "rounds": rounds,
            "order": "native,unified,unified,native",
            "input_parse_timed": false,
            "selector_parse_timed": false,
            "output_serialization_timed": false,
        },
    }))
}

fn parse_usize(value: &str, name: &str) -> Result<usize> {
    let value = value
        .parse::<usize>()
        .map_err(|error| format!("invalid {name} {value:?}: {error}"))?;
    Ok(value)
}

fn parse_args() -> Result<Config> {
    let mut counts = DEFAULT_COUNTS.to_vec();
    let mut warmup = DEFAULT_WARMUP;
    let mut rounds = DEFAULT_ROUNDS;
    let mut args = std::env::args().skip(1);
    while let Some(argument) = args.next() {
        let (name, inline) = argument
            .split_once('=')
            .map_or((argument.as_str(), None), |(name, value)| {
                (name, Some(value))
            });
        match name {
            "--counts" => {
                let value = inline
                    .map(str::to_owned)
                    .or_else(|| args.next())
                    .ok_or("--counts requires a comma-separated value")?;
                counts = value
                    .split(',')
                    .map(|value| parse_usize(value, "count"))
                    .collect::<Result<Vec<_>>>()?;
                if counts.is_empty() || counts.contains(&0) {
                    return Err("--counts requires only positive paragraph counts".into());
                }
            },
            "--warmup" => {
                let value = inline
                    .map(str::to_owned)
                    .or_else(|| args.next())
                    .ok_or("--warmup requires a value")?;
                warmup = parse_usize(&value, "warmup")?;
            },
            "--rounds" => {
                let value = inline
                    .map(str::to_owned)
                    .or_else(|| args.next())
                    .ok_or("--rounds requires a value")?;
                rounds = parse_usize(&value, "rounds")?;
            },
            "--help" | "-h" => {
                println!(
                    "usage: unified-rtf-remove-baseline [--counts N,N] [--warmup N] [--rounds N]"
                );
                std::process::exit(0);
            },
            other => return Err(format!("unknown argument {other:?}").into()),
        }
    }
    if rounds == 0 {
        return Err("--rounds must be positive".into());
    }
    Ok(Config {
        counts,
        warmup,
        rounds,
    })
}

fn env_value(name: &str) -> Option<String> {
    std::env::var(name).ok().filter(|value| !value.is_empty())
}

fn run() -> Result<Value> {
    let config = parse_args()?;
    let mut cases = Vec::with_capacity(config.counts.len() * 2);
    for count in config.counts {
        let source = corpus(count)?;
        for shape in [Shape::SinglePath, Shape::QueryAll] {
            cases.push(measure_case(
                &source,
                count,
                shape,
                config.warmup,
                config.rounds,
            )?);
        }
    }
    Ok(json!({
        "schema": "rtf_remove_benchmark.v1",
        "source_commit": env_value("LITCHI_SOURCE_COMMIT"),
        "source_digest": env_value("LITCHI_SOURCE_DIGEST"),
        "rustc_version": env_value("LITCHI_RUSTC_VERSION"),
        "target": env_value("LITCHI_TARGET"),
        "machine": env_value("LITCHI_MACHINE"),
        "binary_path": env_value("LITCHI_BINARY_PATH"),
        "binary_bytes": env_value("LITCHI_BINARY_BYTES"),
        "binary_sha256": env_value("LITCHI_BINARY_SHA256"),
        "build_command": env_value("LITCHI_BUILD_COMMAND"),
        "cases": cases,
    }))
}

fn main() {
    match run() {
        Ok(report) => {
            println!(
                "{}",
                serde_json::to_string_pretty(&report).expect("benchmark report serializes")
            );
        },
        Err(error) => {
            eprintln!("unified RTF removal benchmark failed: {error}");
            std::process::exit(1);
        },
    }
}
