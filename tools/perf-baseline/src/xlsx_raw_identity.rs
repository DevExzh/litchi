//! Matched raw timeline-cache insertion probes for W4.
//!
//! The probe keeps corpus construction, raw values, and output verification
//! outside the measured operation.  Build the same source against the
//! pre-repair and candidate revisions, pass each exact source revision on the
//! command line, and compare the emitted allocation rows and external
//! Callgrind runs.  The operation itself is the same owned edit, raw identity
//! batch, and commit lifecycle at both revisions.

use crate::allocation_metrics::{self, Sample as AllocationSample};
#[cfg(not(feature = "xlsx-raw-candidate"))]
use litchi_opc::SourceLimits;
use litchi_opc::{BlobPart, PackURI};
use litchi_xlsx::Package;
use litchi_xlsx::timeline::{self, CacheDefinition, FilterType, Range, State};
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::{error::Error, ffi::OsString, fs::OpenOptions, io::Write, path::PathBuf};

type ProbeResult<T> = Result<T, Box<dyn Error>>;

const SCHEMA: &str = "litchi.xlsx.raw-identity-lifecycle.v1";
const CORPUS_GENERATOR: &str = "litchi-xlsx-raw-identity-package-v1";
const LARGE_PART_BYTES: usize = 512 * 1024;
const LARGE_PART_COUNT: usize = 16;
const DEFAULT_ITERATIONS: usize = 5;
const DEFAULT_WARMUPS: usize = 2;

#[derive(Clone, Copy, Debug)]
struct CaseSpec {
    name: &'static str,
    extra_parts: usize,
    extra_part_bytes: usize,
    cache_count: usize,
}

const CASES: [CaseSpec; 4] = [
    CaseSpec {
        name: "small_single",
        extra_parts: 0,
        extra_part_bytes: 0,
        cache_count: 1,
    },
    CaseSpec {
        name: "large_multipart_single",
        extra_parts: LARGE_PART_COUNT,
        extra_part_bytes: LARGE_PART_BYTES,
        cache_count: 1,
    },
    CaseSpec {
        name: "small_multi_cache",
        extra_parts: 0,
        extra_part_bytes: 0,
        cache_count: 8,
    },
    CaseSpec {
        name: "large_multipart_multi_cache",
        extra_parts: LARGE_PART_COUNT,
        extra_part_bytes: LARGE_PART_BYTES,
        cache_count: 8,
    },
];

#[derive(Debug, Serialize)]
struct BinaryIdentity {
    path: String,
    bytes: usize,
    sha256: String,
}

#[derive(Debug, Serialize)]
struct CaseReport {
    name: &'static str,
    extra_parts: usize,
    extra_part_bytes: usize,
    cache_count: usize,
    base_bytes: usize,
    base_sha256: String,
    output_bytes: usize,
    output_sha256: String,
    identity_check: bool,
    allocation_samples: Vec<AllocationSample>,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    api_variant: String,
    source_revision: String,
    corpus_generator: &'static str,
    operation: &'static str,
    allocation_scope: &'static str,
    callgrind_scope: &'static str,
    iterations: usize,
    warmups: usize,
    binary: BinaryIdentity,
    cases: Vec<CaseReport>,
}

fn sha256_hex(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

fn binary_identity() -> ProbeResult<BinaryIdentity> {
    let path = std::env::current_exe()?;
    let bytes = std::fs::read(&path)?;
    Ok(BinaryIdentity {
        path: path.display().to_string(),
        bytes: bytes.len(),
        sha256: sha256_hex(&bytes),
    })
}

fn build_package(spec: CaseSpec) -> ProbeResult<Package> {
    let mut package = Package::new()?.into_plain_opc();
    for index in 0..spec.extra_parts {
        let part_name = PackURI::new(format!("/xl/rawProbe/part-{index:03}.bin"))?;
        let payload = (0..spec.extra_part_bytes)
            .map(|offset| ((offset + index) % 251) as u8)
            .collect();
        package.add_part(Box::new(BlobPart::new(
            part_name,
            "application/octet-stream".to_owned(),
            payload,
        )));
    }
    Ok(Package::from_opc(package)?)
}

fn cache_definition(index: usize) -> ProbeResult<CacheDefinition> {
    Ok(CacheDefinition {
        name: format!("RawProbeCache{index}"),
        uid: None,
        source_name: "Date".to_owned(),
        pivot_tables: Vec::new(),
        state: State {
            selection: Some(Range::new("2026-01-01T00:00:00Z", "2026-01-31T23:59:59Z")?),
            bounds: Range::new("2026-01-01T00:00:00Z", "2026-12-31T23:59:59Z")?,
            extension_list: None,
            single_range_filter_state: Some(true),
            minimal_refresh_version: 0,
            last_refresh_version: 1,
            pivot_cache_id: 1,
            filter_type: FilterType::DateBetween,
        },
        timeline_pivot_filter: None,
        extension_list: None,
    })
}

fn raw_values(spec: CaseSpec) -> ProbeResult<Vec<timeline::raw::Cache>> {
    (0..spec.cache_count)
        .map(|index| {
            Ok(timeline::raw::Cache {
                relationship_id: format!("rIdRawProbe{index}"),
                part_name: format!("/xl/timelineCaches/raw-probe-{index}.xml"),
                definition: cache_definition(index)?,
            })
        })
        .collect()
}

fn insert(source: &Package, values: &[timeline::raw::Cache]) -> ProbeResult<Package> {
    let mut edit = source.edit_timelines()?;
    #[cfg(feature = "xlsx-raw-candidate")]
    edit.insert_caches(values)?;
    #[cfg(not(feature = "xlsx-raw-candidate"))]
    edit.insert_caches(values, &SourceLimits::default())?;
    Ok(edit.commit()?.into_package())
}

fn verify_identity(candidate: &Package, spec: CaseSpec) -> ProbeResult<bool> {
    let package = candidate.clone().into_plain_opc();
    let workbook = package.main_document_part()?.partname().clone();
    let caches = timeline::load_caches(&package, &workbook)?;
    if caches.len() != spec.cache_count {
        return Ok(false);
    }
    Ok(caches.iter().enumerate().all(|(index, cache)| {
        cache.relationship_id == format!("rIdRawProbe{index}")
            && cache.part_name == format!("/xl/timelineCaches/raw-probe-{index}.xml")
    }))
}

fn run_case(spec: CaseSpec, iterations: usize, warmups: usize) -> ProbeResult<CaseReport> {
    let source = build_package(spec)?;
    let values = raw_values(spec)?;
    let base = source.to_bytes()?;

    let reference = insert(&source, &values)?;
    let output = reference.to_bytes()?;
    if !verify_identity(&reference, spec)? {
        return Err(format!("identity verification failed for {}", spec.name).into());
    }
    let output_sha256 = sha256_hex(&output);

    for _ in 0..warmups {
        drop(insert(&source, &values)?);
    }

    let mut allocation_samples = Vec::with_capacity(iterations);
    for _ in 0..iterations {
        let region = allocation_metrics::begin();
        let candidate = insert(&source, &values)?;
        let sample = region.finish().ok_or("allocator region unavailable")?;
        let candidate_bytes = candidate.to_bytes()?;
        if sha256_hex(&candidate_bytes) != output_sha256 || !verify_identity(&candidate, spec)? {
            return Err(format!("candidate verification failed for {}", spec.name).into());
        }
        allocation_samples.push(sample);
    }

    Ok(CaseReport {
        name: spec.name,
        extra_parts: spec.extra_parts,
        extra_part_bytes: spec.extra_part_bytes,
        cache_count: spec.cache_count,
        base_bytes: base.len(),
        base_sha256: sha256_hex(&base),
        output_bytes: output.len(),
        output_sha256,
        identity_check: true,
        allocation_samples,
    })
}

fn parse_positive(value: OsString, flag: &str) -> ProbeResult<usize> {
    let value = value
        .into_string()
        .map_err(|_| format!("{flag} must be UTF-8"))?;
    let parsed = value
        .parse::<usize>()
        .map_err(|error| format!("invalid {flag} {value:?}: {error}"))?;
    if parsed == 0 {
        return Err(format!("{flag} must be positive").into());
    }
    Ok(parsed)
}

fn next_value(values: &mut impl Iterator<Item = OsString>, flag: &str) -> ProbeResult<OsString> {
    values
        .next()
        .ok_or_else(|| format!("missing value for {flag}").into())
}

/// Run the allocator probe. `--source-revision` is required so a report
/// cannot be mistaken for a measurement from an unrecorded checkout.
pub fn run_from_args(args: impl IntoIterator<Item = OsString>) -> ProbeResult<()> {
    let mut iterations = DEFAULT_ITERATIONS;
    let mut warmups = DEFAULT_WARMUPS;
    let mut api_variant = None;
    let mut source_revision = None;
    let mut selected_case = None;
    let mut json_path = None::<PathBuf>;
    let mut values = args.into_iter();
    while let Some(argument) = values.next() {
        match argument.to_string_lossy().as_ref() {
            "--iterations" => {
                iterations =
                    parse_positive(next_value(&mut values, "--iterations")?, "--iterations")?
            },
            "--warmups" => {
                warmups = parse_positive(next_value(&mut values, "--warmups")?, "--warmups")?
            },
            "--api-variant" => {
                api_variant = Some(
                    next_value(&mut values, "--api-variant")?
                        .into_string()
                        .map_err(|_| "--api-variant must be UTF-8")?,
                )
            },
            "--source-revision" => {
                source_revision = Some(
                    next_value(&mut values, "--source-revision")?
                        .into_string()
                        .map_err(|_| "--source-revision must be UTF-8")?,
                )
            },
            "--case" => {
                selected_case = Some(
                    next_value(&mut values, "--case")?
                        .into_string()
                        .map_err(|_| "--case must be UTF-8")?,
                )
            },
            "--json" => json_path = Some(PathBuf::from(next_value(&mut values, "--json")?)),
            "--help" | "-h" => {
                println!(
                    "xlsx-raw-identity --api-variant <original|candidate> --source-revision <commit> [--case <name>] [--iterations N] [--warmups N] [--json PATH]"
                );
                return Ok(());
            },
            unknown => return Err(format!("unknown argument {unknown}").into()),
        }
    }
    let api_variant = api_variant.ok_or("--api-variant is required")?;
    if api_variant != "original" && api_variant != "candidate" {
        return Err("--api-variant must be original or candidate".into());
    }
    let source_revision = source_revision.ok_or("--source-revision is required")?;
    if source_revision.len() != 40 || !source_revision.bytes().all(|byte| byte.is_ascii_hexdigit())
    {
        return Err("--source-revision must be a 40-character hexadecimal commit".into());
    }

    let cases = CASES
        .into_iter()
        .filter(|case| {
            selected_case
                .as_deref()
                .is_none_or(|selected| selected == case.name)
        })
        .collect::<Vec<_>>();
    if cases.is_empty() {
        return Err("--case does not name a known raw identity case".into());
    }
    let report = Report {
        schema: SCHEMA,
        api_variant,
        source_revision,
        corpus_generator: CORPUS_GENERATOR,
        operation: "owned Package::edit_timelines -> raw::Edit::insert_caches -> commit",
        allocation_scope: "allocator region surrounds edit creation, raw batch admission, commit, and candidate construction; output serialization, identity verification, and drops are outside the region",
        callgrind_scope: "run the same binary and case under callgrind with identical iterations and warmups; compare Ir by exact source_revision and binary SHA-256",
        iterations,
        warmups,
        binary: binary_identity()?,
        cases: cases
            .into_iter()
            .map(|case| run_case(case, iterations, warmups))
            .collect::<ProbeResult<Vec<_>>>()?,
    };
    let encoded = serde_json::to_vec_pretty(&report)?;
    if let Some(path) = json_path {
        let mut file = OpenOptions::new().write(true).create_new(true).open(path)?;
        file.write_all(&encoded)?;
        file.write_all(b"\n")?;
        file.flush()?;
    } else {
        let mut stdout = std::io::stdout().lock();
        stdout.write_all(&encoded)?;
        stdout.write_all(b"\n")?;
    }
    Ok(())
}
