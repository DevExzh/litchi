//! Paired XLSB creation: the native package constructor plus ready workbook
//! projection versus the unified creation entry point.
//!
//! The timed operation is deliberately small.  The native leg measures
//! `Package::new().into_workbook()` and the unified leg measures
//! `litchi_ops::create(FileFormat::Xlsb)`.  Serialization, caller-readiness
//! checks, ZIP equivalence, and destruction are outside the clocks.  This is
//! creation-only evidence; it makes no claim about opening or writing a file.

use litchi_core::FileFormat;
use litchi_core::sheet::{WorkbookTrait, Worksheet as _};
use litchi_ops::File;
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use soapberry_zip::office::ArchiveReader;
use std::collections::{BTreeMap, BTreeSet};
use std::error::Error;
use std::hint::black_box;
use std::time::Instant;

type Result<T> = std::result::Result<T, Box<dyn Error + Send + Sync>>;

const DEFAULT_ROUNDS: usize = 9;
const SHEET_NAME: &str = "Sheet1";

struct Config {
    rounds: usize,
}

fn native_create() -> Result<litchi_xlsb::Package> {
    Ok(litchi_xlsb::Package::new()?)
}

fn unified_create() -> Result<File> {
    Ok(litchi_ops::create(FileFormat::Xlsb)?)
}

fn verify_native_ready(package: &litchi_xlsb::Package) -> Result<()> {
    let workbook = package.workbook()?;
    let names = WorkbookTrait::worksheet_names(workbook);
    if names.len() != 1 || names.first().map(String::as_str) != Some(SHEET_NAME) {
        return Err(format!("native creation returned worksheet names {names:?}").into());
    }
    let sheet = WorkbookTrait::worksheet_by_name(workbook, SHEET_NAME)?;
    if sheet.name() != SHEET_NAME {
        return Err(format!("native Sheet1 lookup returned worksheet {:?}", sheet.name()).into());
    }
    Ok(())
}

fn verify_unified_ready(file: &File) -> Result<()> {
    if file.format() != Some(FileFormat::Xlsb) {
        return Err(format!(
            "unified creation returned format {:?}, expected XLSB",
            file.format()
        )
        .into());
    }
    let workbook = file.workbook()?;
    let names = workbook.worksheet_names()?;
    if names.len() != 1 || names.first().map(String::as_str) != Some(SHEET_NAME) {
        return Err(format!("unified creation returned worksheet names {names:?}").into());
    }
    let sheet = WorkbookTrait::worksheet_by_name(&workbook, SHEET_NAME)?;
    if sheet.name() != SHEET_NAME {
        return Err(format!(
            "unified Sheet1 lookup returned worksheet {:?}",
            sheet.name()
        )
        .into());
    }
    Ok(())
}

fn native_bytes(package: litchi_xlsb::Package) -> Result<Vec<u8>> {
    Ok(package.to_bytes()?)
}

fn unified_bytes(file: &File) -> Result<Vec<u8>> {
    Ok(file.to_bytes()?)
}

fn sha256(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

fn decoded_members(bytes: &[u8]) -> Result<BTreeMap<String, Vec<u8>>> {
    let archive = ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default())?;
    let mut members = BTreeMap::new();
    for name in archive.file_names() {
        members.insert(name.to_owned(), archive.read(name)?);
    }
    Ok(members)
}

fn member_description(bytes: Option<&Vec<u8>>) -> Value {
    match bytes {
        Some(bytes) => json!({
            "present": true,
            "bytes": bytes.len(),
            "sha256": sha256(bytes),
        }),
        None => json!({"present": false}),
    }
}

/// Prefer exact bytes, then compare every decoded ZIP member.  A decoded
/// match is useful when ZIP container metadata differs while the package
/// payload is identical; a payload mismatch remains a hard failure.
fn verify_equivalence(native: &[u8], unified: &[u8]) -> Result<Value> {
    let native_sha256 = sha256(native);
    let unified_sha256 = sha256(unified);
    if native == unified {
        return Ok(json!({
            "kind": "exact_bytes",
            "native_bytes": native.len(),
            "unified_bytes": unified.len(),
            "native_sha256": native_sha256,
            "unified_sha256": unified_sha256,
            "differing_members": [],
        }));
    }

    let native_members = decoded_members(native)?;
    let unified_members = decoded_members(unified)?;
    let names = native_members
        .keys()
        .chain(unified_members.keys())
        .cloned()
        .collect::<BTreeSet<_>>();
    let differing_members = names
        .into_iter()
        .filter_map(|name| {
            let left = native_members.get(&name);
            let right = unified_members.get(&name);
            (left != right).then(|| {
                json!({
                    "member": name,
                    "native": member_description(left),
                    "unified": member_description(right),
                })
            })
        })
        .collect::<Vec<_>>();
    if differing_members.is_empty() {
        Ok(json!({
            "kind": "decoded_member_bytes",
            "native_bytes": native.len(),
            "unified_bytes": unified.len(),
            "native_sha256": native_sha256,
            "unified_sha256": unified_sha256,
            "decoded_member_count": native_members.len(),
            "differing_members": [],
        }))
    } else {
        Err(format!(
            "native and unified XLSB creation outputs differ in decoded ZIP members: {}",
            serde_json::to_string(&differing_members)?
        )
        .into())
    }
}

fn median_upper_middle(values: &[u128]) -> u128 {
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    sorted[sorted.len() / 2]
}

fn measure(rounds: usize) -> Result<Value> {
    // The warm-up pays all correctness and publication checks before the
    // clocks start.  The native and unified outputs must be equivalent, and
    // both freshly created owners must already expose Sheet1 to callers.
    let native_warm = native_create()?;
    verify_native_ready(&native_warm)?;
    let native_output = native_bytes(native_warm)?;
    let unified_warm = unified_create()?;
    verify_unified_ready(&unified_warm)?;
    let unified_output = unified_bytes(&unified_warm)?;
    let warm_equivalence = verify_equivalence(&native_output, &unified_output)?;

    let mut native_ns = Vec::with_capacity(rounds * 2);
    let mut unified_ns = Vec::with_capacity(rounds * 2);
    for _ in 0..rounds {
        // ABBA keeps adjacent samples of each route paired while retaining
        // the requested 2 * rounds samples per leg.
        for native_leg in [true, false, false, true] {
            if native_leg {
                let start = Instant::now();
                let workbook = native_create()?;
                let elapsed = start.elapsed().as_nanos();
                black_box(&workbook);
                native_ns.push(elapsed);
                drop(workbook);
            } else {
                let start = Instant::now();
                let file = unified_create()?;
                let elapsed = start.elapsed().as_nanos();
                black_box(&file);
                unified_ns.push(elapsed);
                drop(file);
            }
        }
    }

    let native_median_ns = median_upper_middle(&native_ns);
    let unified_median_ns = median_upper_middle(&unified_ns);
    Ok(json!({
        "native_route": "xlsb_native_package_new_into_workbook_v1",
        "unified_route": "xlsb_unified_create_xlsb_v1",
        "native_ready": true,
        "unified_ready": true,
        "native_output_bytes": native_output.len(),
        "unified_output_bytes": unified_output.len(),
        "native_output_sha256": sha256(&native_output),
        "unified_output_sha256": sha256(&unified_output),
        "warm_equivalence": warm_equivalence,
        "native_ns": native_ns,
        "unified_ns": unified_ns,
        "native_median_ns": native_median_ns,
        "unified_median_ns": unified_median_ns,
        "unified_over_native": unified_median_ns as f64 / native_median_ns as f64,
    }))
}

fn parse_args() -> Result<Config> {
    let mut rounds = DEFAULT_ROUNDS;
    let mut args = std::env::args().skip(1);
    while let Some(argument) = args.next() {
        let (key, inline) = argument
            .split_once('=')
            .map_or((argument.as_str(), None), |(key, value)| (key, Some(value)));
        match key {
            "--help" | "-h" => {
                println!("usage: unified-xlsb-create-baseline [--rounds N]");
                std::process::exit(0);
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
    Ok(Config { rounds })
}

fn build_provenance() -> Value {
    json!({
        "crate": env!("CARGO_PKG_NAME"),
        "crate_version": env!("CARGO_PKG_VERSION"),
        "rustc": std::env::var("LITCHI_RUSTC_VERSION").unwrap_or_else(|_| "unspecified".to_owned()),
        "target": std::env::var("LITCHI_TARGET").unwrap_or_else(|_| "unspecified".to_owned()),
        "source_commit": std::env::var("LITCHI_SOURCE_COMMIT").unwrap_or_else(|_| "unspecified".to_owned()),
        "source_digest": std::env::var("LITCHI_SOURCE_DIGEST").unwrap_or_else(|_| "unspecified".to_owned()),
        "binary_sha256": std::env::var("LITCHI_BINARY_SHA256").unwrap_or_else(|_| "unspecified".to_owned()),
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
    match measure(config.rounds) {
        Ok(result) => {
            eprintln!("measured XLSB creation");
            results.push(result);
        },
        Err(error) => {
            eprintln!("failed XLSB creation benchmark: {error}");
            failures.push(json!({
                "route": "native_package_new_into_workbook_vs_unified_create_xlsb",
                "phase": "creation, caller-readiness, serialization, equivalence, or timing",
                "error": error.to_string(),
            }));
        },
    }

    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "schema": 1,
            "format": "xlsb",
            "workload": "minimal-xlsb-creation",
            "rounds": config.rounds,
            "order": "ABBA",
            "samples_per_leg": config.rounds * 2,
            "routes": [
                "xlsb_native_package_new_into_workbook_v1",
                "xlsb_unified_create_xlsb_v1",
            ],
            "timing_scope": "native Package::new().into_workbook() versus unified litchi_ops::create(FileFormat::Xlsb); only initial package construction and ready workbook or File projection are timed",
            "excluded_from_timing": [
                "caller-readiness checks for Sheet1",
                "native and unified output serialization",
                "ZIP equivalence and decoded-member checks",
                "warm-up and destruction",
                "opening or writing a caller-provided file",
            ],
            "caller_readiness": "each warm-up owner exposes exactly one worksheet named Sheet1 through the native or unified workbook view and resolves Sheet1 by name",
            "equivalence_scope": "warm-up outputs are compared by exact ZIP bytes, falling back to decoded ZIP member bytes with per-member length and SHA-256 reporting",
            "build_provenance": build_provenance(),
            "results": results,
            "failure_matrix_scope": [
                "native Package::new plus into_workbook",
                "unified create(FileFormat::Xlsb)",
                "native Sheet1 caller readiness",
                "unified Sheet1 caller readiness",
                "native and unified serialization",
                "exact or decoded-member equivalence",
                "ABBA timed samples",
            ],
            "failure_matrix": failures,
            "performance_claim": "creation-only timing evidence; no open or write budget is claimed",
        }))?
    );
    if failures.is_empty() {
        Ok(())
    } else {
        Err("creation benchmark contains failures; completed samples and failures are recorded in JSON".into())
    }
}
