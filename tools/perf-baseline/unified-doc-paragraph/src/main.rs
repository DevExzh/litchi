//! Paired legacy DOC paragraph replacement: the native body-text transaction
//! versus one unified `File::set` call.
//!
//! The corpus, owner opening, target preparation, warm-up, output validation,
//! and output serialization are outside the clocks except where stated in the
//! emitted timing scope.  The native leg opens a body-text snapshot from a
//! prepared byte buffer inside every timed sample.  The unified leg opens a
//! a freshly opened retained source owner before every unified timed sample;
//! `File::set` materializes the source-backed DOC owner into the variable-width
//! body-text editor inside that sample.  Fresh owners keep the facade's
//! OnceLock snapshot cache from warming later samples.  This keeps the
//! source-promotion boundary visible on both routes while keeping caller
//! parsing and serialization out of timing.

use litchi_core::Position;
use litchi_doc::Builder;
use litchi_doc::body_text::{Projection, Snapshot};
use litchi_doc::tracked_revision::Limits;
use litchi_ops::{File, OpenOptions, SourceMode};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use std::error::Error;
use std::hint::black_box;
use std::time::Instant;

type Result<T> = std::result::Result<T, Box<dyn Error + Send + Sync>>;

const DEFAULT_COUNTS: [usize; 4] = [1, 10, 100, 1_000];
const DEFAULT_ROUNDS: usize = 9;

struct Config {
    counts: Vec<usize>,
    rounds: usize,
}

fn paragraph_text(position: usize) -> String {
    format!("seed paragraph {position:05}")
}

fn replacement_text(position: usize) -> String {
    format!("updated paragraph {position:05} with a deliberately longer payload")
}

fn corpus(count: usize) -> Result<Vec<u8>> {
    if count == 0 {
        return Err("the paragraph corpus must contain at least one paragraph".into());
    }
    let mut builder = Builder::new();
    for position in 0..count {
        builder.add_paragraph(&paragraph_text(position))?;
    }
    let mut package = builder.build()?;
    Ok(package.to_bytes()?)
}

fn open_unified(bytes: &[u8]) -> Result<File> {
    let options = OpenOptions::new().source_mode(SourceMode::Auto);
    Ok(litchi_ops::open_bytes_with(bytes.to_vec(), &options)?)
}

/// This boundary deliberately includes body-text source promotion and native
/// publication.  `bytes.to_vec()` is prepared by the caller and cloned here,
/// matching the unified source's materialization into an owned body snapshot.
fn native_once(bytes: &[u8], position: usize, replacement: &str) -> Result<Snapshot> {
    let snapshot = Snapshot::open(bytes.to_vec(), Limits::default())?;
    let mut edit = snapshot.edit()?;
    edit.replace_paragraph(Position::new(position), replacement)?;
    let commit = edit.commit()?;
    Ok(commit.into_parts().0)
}

/// The path is parsed once outside timing.  `&Path` is accepted by the
/// generic target conversion without reparsing the selector in this call.
fn unified_once(
    file: &File,
    path: &litchi_core::selector::Path<'static>,
    replacement: &str,
) -> Result<File> {
    Ok(file.set(path, litchi_ops::props! { "text" => replacement })?)
}

fn sha256(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

fn text_digest(texts: &[String]) -> String {
    sha256(texts.join("\n").as_bytes())
}

fn paragraph_texts(bytes: &[u8]) -> Result<Vec<String>> {
    let snapshot = Snapshot::open(bytes.to_vec(), Limits::default())?;
    Ok(snapshot
        .paragraphs(Projection::All)?
        .into_iter()
        .map(|paragraph| paragraph.text().to_owned())
        .collect())
}

fn verify_original(bytes: &[u8], count: usize) -> Result<Vec<String>> {
    let texts = paragraph_texts(bytes)?;
    let expected = (0..count).map(paragraph_text).collect::<Vec<_>>();
    if texts != expected {
        return Err(format!(
            "generated DOC paragraph text differs from the expected corpus: observed count {}, expected count {count}",
            texts.len()
        )
        .into());
    }
    Ok(texts)
}

fn verify_output(
    bytes: &[u8],
    count: usize,
    position: usize,
    replacement: &str,
) -> Result<Vec<String>> {
    let mut expected = (0..count).map(paragraph_text).collect::<Vec<_>>();
    expected[position] = replacement.to_owned();
    let observed = paragraph_texts(bytes)?;
    if observed != expected {
        return Err(format!(
            "DOC paragraph output mismatch at position {position}: observed count {}, expected count {count}",
            observed.len()
        )
        .into());
    }
    if bytes.is_empty() {
        return Err("DOC paragraph publication returned an empty compound file".into());
    }
    Ok(observed)
}

fn verify_exact(native: &[u8], unified: &[u8]) -> Result<Value> {
    let native_sha256 = sha256(native);
    let unified_sha256 = sha256(unified);
    if native != unified {
        return Err(format!(
            "native and unified DOC paragraph outputs differ: native {} bytes/{native_sha256}, unified {} bytes/{unified_sha256}",
            native.len(),
            unified.len()
        )
        .into());
    }
    Ok(json!({
        "kind": "exact_bytes",
        "native_bytes": native.len(),
        "unified_bytes": unified.len(),
        "native_sha256": native_sha256,
        "unified_sha256": unified_sha256,
    }))
}

fn median(values: &[u128]) -> u128 {
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    sorted[sorted.len() / 2]
}

fn measure(bytes: &[u8], count: usize, rounds: usize) -> Result<Value> {
    let position = count / 2;
    let replacement = replacement_text(position);
    let path_text = format!("/body/p[{}]", position + 1);
    let path = litchi_core::selector::Path::parse(&path_text)?.into_owned();

    let original_texts = verify_original(bytes, count)?;
    let native_warm = native_once(bytes, position, &replacement)?;
    let unified_warm = {
        let unified = open_unified(bytes)?;
        unified_once(&unified, &path, &replacement)?
    };
    let native_warm_bytes = native_warm.bytes();
    let unified_warm_bytes = unified_warm.to_bytes()?;
    let native_warm_texts = verify_output(native_warm_bytes, count, position, &replacement)?;
    let unified_warm_texts = verify_output(&unified_warm_bytes, count, position, &replacement)?;
    if native_warm_texts != unified_warm_texts {
        return Err("native and unified DOC paragraph text projections differ".into());
    }
    let warm_equivalence = verify_exact(native_warm_bytes, &unified_warm_bytes)?;
    let native_output_sha256 = sha256(native_warm_bytes);
    let unified_output_sha256 = sha256(&unified_warm_bytes);
    if native_warm_bytes == bytes || unified_warm_bytes == bytes {
        return Err("the warm-up paragraph replacement produced unchanged bytes".into());
    }
    drop((native_warm, unified_warm));

    let mut native_ns = Vec::with_capacity(rounds * 2);
    let mut unified_ns = Vec::with_capacity(rounds * 2);
    for _ in 0..rounds {
        // ABBA: native, unified, unified, native.
        for native_leg in [true, false, false, true] {
            if native_leg {
                let start = Instant::now();
                let output = native_once(black_box(bytes), position, black_box(&replacement))?;
                native_ns.push(start.elapsed().as_nanos());
                black_box(&output);
                drop(output);
            } else {
                // Open a fresh source-backed owner before every sample.  The
                // DOC facade caches its promoted Snapshot in an OnceLock on
                // the File, so reusing the warm-up owner would time a warm
                // cache rather than the source-promotion boundary paired with
                // native Snapshot::open below.
                let unified = open_unified(bytes)?;
                let start = Instant::now();
                let output = unified_once(
                    black_box(&unified),
                    black_box(&path),
                    black_box(&replacement),
                )?;
                unified_ns.push(start.elapsed().as_nanos());
                black_box(&output);
                drop(output);
            }
        }
    }

    let native_median_ns = median(&native_ns);
    let unified_median_ns = median(&unified_ns);
    Ok(json!({
        "paragraphs": count,
        "operation": "paragraph_text_set",
        "target_position": position,
        "target_path": path_text,
        "replacement_utf16_units": replacement.encode_utf16().count(),
        "input_paragraph_text_digest": text_digest(&original_texts),
        "output_paragraph_text_digest": text_digest(&native_warm_texts),
        "native_route": "doc_source_promoted_native_body_text_snapshot_set_v1",
        "unified_route": "doc_source_promoted_unified_file_set_fresh_source_v2",
        "warm_equivalence": warm_equivalence,
        "native_output_sha256": native_output_sha256,
        "unified_output_sha256": unified_output_sha256,
        "native_ns": native_ns,
        "unified_ns": unified_ns,
        "native_median_ns": native_median_ns,
        "unified_median_ns": unified_median_ns,
        "unified_over_native": unified_median_ns as f64 / native_median_ns as f64,
        "input_cost": {
            "native": {
                "source_bytes_prepared_outside_timing": true,
                "snapshot_open_inside_timing": true,
                "timed_call_boundary": "Snapshot::open(bytes.clone()) + Edit::replace_paragraph + Edit::commit; committed Snapshot bytes are handed off"
            },
            "unified": {
                "file_open_outside_timing": true,
                "fresh_file_per_sample": true,
                "path_prepared_outside_timing": true,
                "source_promotion_inside_timing": true,
                "timed_call_boundary": "File::set(&parsed_path, text property); returned File publication is not serialized inside timing"
            },
            "pairing": "one freshly opened source-backed owner and one variable-width paragraph replacement per sample; no path parsing or output verification occurs inside timing"
        }
    }))
}

fn parse_counts(value: &str) -> Result<Vec<usize>> {
    let counts = value
        .split(',')
        .map(|part| {
            let count = part.parse::<usize>()?;
            if count == 0 {
                return Err("counts must be positive".into());
            }
            Ok(count)
        })
        .collect::<Result<Vec<_>>>()?;
    if counts.is_empty() {
        return Err("--counts requires at least one count".into());
    }
    Ok(counts)
}

fn parse_args() -> Result<Config> {
    let mut counts = DEFAULT_COUNTS.to_vec();
    let mut rounds = DEFAULT_ROUNDS;
    let mut args = std::env::args().skip(1);
    while let Some(argument) = args.next() {
        let (key, inline) = argument
            .split_once('=')
            .map_or((argument.as_str(), None), |(key, value)| (key, Some(value)));
        match key {
            "--help" | "-h" => {
                println!("usage: unified-doc-paragraph-baseline [--counts N[,N...]] [--rounds N]");
                std::process::exit(0);
            },
            "--counts" => {
                let value = inline
                    .map(str::to_owned)
                    .or_else(|| args.next())
                    .ok_or("--counts requires a comma-separated value")?;
                counts = parse_counts(&value)?;
            },
            "--rounds" => {
                let value = inline
                    .map(str::to_owned)
                    .or_else(|| args.next())
                    .ok_or("--rounds requires a value")?;
                rounds = value.parse()?;
            },
            _ => return Err(format!("unknown argument {argument:?}").into()),
        }
    }
    if rounds < 3 {
        return Err("at least three ABBA rounds are required".into());
    }
    Ok(Config { counts, rounds })
}

fn build_provenance() -> Value {
    let source_manifest = std::env::var("LITCHI_SOURCE_MANIFEST")
        .ok()
        .and_then(|value| serde_json::from_str(&value).ok())
        .unwrap_or_else(|| json!("unspecified"));
    json!({
        "crate": env!("CARGO_PKG_NAME"),
        "crate_version": env!("CARGO_PKG_VERSION"),
        "rustc": std::env::var("LITCHI_RUSTC_VERSION").unwrap_or_else(|_| "unspecified".to_owned()),
        "target": std::env::var("LITCHI_TARGET").unwrap_or_else(|_| "unspecified".to_owned()),
        "source_commit": std::env::var("LITCHI_SOURCE_COMMIT").unwrap_or_else(|_| "unspecified".to_owned()),
        "source_digest": std::env::var("LITCHI_SOURCE_DIGEST").unwrap_or_else(|_| "unspecified".to_owned()),
        "source_manifest": source_manifest,
        "binary_manifest": {
            "path": std::env::var("LITCHI_BINARY_PATH").unwrap_or_else(|_| "unspecified".to_owned()),
            "bytes": std::env::var("LITCHI_BINARY_BYTES").unwrap_or_else(|_| "unspecified".to_owned()),
            "sha256": std::env::var("LITCHI_BINARY_SHA256").unwrap_or_else(|_| "unspecified".to_owned()),
        },
        "build_command": std::env::var("LITCHI_BUILD_COMMAND").unwrap_or_else(|_| "unspecified".to_owned()),
        "machine": {
            "os": std::env::consts::OS,
            "arch": std::env::consts::ARCH,
            "parallelism": std::thread::available_parallelism().map_or(0, |value| value.get()),
            "label": std::env::var("LITCHI_MACHINE").unwrap_or_else(|_| "unspecified".to_owned()),
        },
    })
}

fn main() -> Result<()> {
    let config = parse_args()?;
    let mut results = Vec::new();
    let mut failures = Vec::new();
    for count in config.counts.iter().copied() {
        let bytes = match corpus(count) {
            Ok(bytes) => bytes,
            Err(error) => {
                eprintln!("failed DOC corpus generation: {count} paragraphs: {error}");
                failures.push(json!({
                    "paragraphs": count,
                    "phase": "corpus",
                    "error": error.to_string(),
                }));
                continue;
            },
        };
        let artifact = json!({
            "generator": "litchi_doc::Builder::build()?.to_bytes(); one plain seed paragraph per requested position",
            "paragraphs": count,
            "bytes": bytes.len(),
            "sha256": sha256(&bytes),
        });
        match measure(&bytes, count, config.rounds) {
            Ok(mut result) => {
                result["artifact"] = artifact;
                eprintln!("measured DOC paragraph replacement: {count} paragraphs");
                results.push(result);
            },
            Err(error) => {
                eprintln!("failed DOC paragraph replacement: {count} paragraphs: {error}");
                failures.push(json!({
                    "paragraphs": count,
                    "phase": "open, set, publication, validation, or timing",
                    "artifact": artifact,
                    "error": error.to_string(),
                }));
            },
        }
    }

    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "schema": 1,
            "format": "doc",
            "workload": "one-variable-width-body-paragraph-set",
            "counts": config.counts,
            "rounds": config.rounds,
            "order": "ABBA",
            "samples_per_leg": config.rounds * 2,
            "routes": [
                "doc_source_promoted_native_body_text_snapshot_set_v1",
                "doc_source_promoted_unified_file_set_fresh_source_v2",
            ],
            "timing_scope": "native Snapshot::open(bytes.clone()) + one variable-width Edit::replace_paragraph + commit versus one File::set on a freshly opened source-backed File with a prepared paragraph path; source promotion is inside both legs, while fresh File open, output serialization, and verification are outside",
            "excluded_from_timing": [
                "corpus generation and source fingerprint manifest",
                "fresh unified File open before every unified sample",
                "path parsing and target preparation",
                "warm-up",
                "output serialization for the unified File",
                "paragraph text and exact-byte validation",
                "output destruction",
            ],
            "equivalence_scope": "warm-up outputs must match byte-for-byte and paragraph-for-paragraph; every timed leg uses the same fixed corpus and replacement",
            "build_provenance": build_provenance(),
            "results": results,
            "failure_matrix_scope": [
                "Builder corpus generation",
                "native body-text snapshot open and replacement",
                "unified source-backed File open and set",
                "native and unified output text",
                "native and unified exact bytes",
                "ABBA timed samples",
            ],
            "failure_matrix": failures,
            "performance_claim": "scoped single-paragraph timing evidence only; no batch scaling, allocation, RSS, open, save, or full-goal acceptance claim",
        }))?
    );
    if failures.is_empty() {
        Ok(())
    } else {
        Err("DOC paragraph benchmark contains failures; completed samples and failures are recorded in JSON".into())
    }
}
