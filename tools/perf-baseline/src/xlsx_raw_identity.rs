//! Matched XLSX raw-identity lifecycle probes for W4.
//!
//! This harness uses one API on both the immediate pre-repair revision
//! (`0befb112d96d2409e6f40c9abfeb4813665905ad`) and candidate
//! (`5726f03c7db2dd7dd539000b45f109f14ea8c33a`). It covers owned lazy open,
//! first selected read, raw slicer/timeline insert, typed replacement, commit,
//! replay, inverse, stale-patch refusal, no-op, native identity, untouched
//! bytes, and deferred decode counters.

use crate::allocation_metrics::{self, Sample as AllocationSample};
use litchi_opc::{BlobPart, OpcPackage, PackURI, SourceLimits};
use litchi_xlsx::Package;
use litchi_xlsx::slicer;
use litchi_xlsx::timeline::{self, CacheDefinition, FilterType, Range, State};
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::{error::Error, ffi::OsString, fs::OpenOptions, io::Write, path::PathBuf, time::Instant};

type ProbeResult<T> = Result<T, Box<dyn Error>>;

const SCHEMA: &str = "litchi.xlsx.raw-identity-lifecycle.v2";
const CORPUS_GENERATOR: &str = "litchi-xlsx-raw-identity-package-v2";
const LARGE_PART_BYTES: usize = 512 * 1024;
const LARGE_PART_COUNT: usize = 16;
const DEFAULT_ITERATIONS: usize = 5;
const DEFAULT_WARMUPS: usize = 2;
const WORKBOOK: &str = "/xl/workbook.xml";
const WORKSHEET: &str = "/xl/worksheets/sheet1.xml";
const EXISTING_NAME: &str = "RawProbeExisting";

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

#[derive(Clone, Copy, Debug)]
enum Operation {
    SlicerInsert,
    TimelineInsert,
    SlicerReplace,
    TimelineReplace,
}

impl Operation {
    const ALL: [Self; 4] = [
        Self::SlicerInsert,
        Self::TimelineInsert,
        Self::SlicerReplace,
        Self::TimelineReplace,
    ];

    const fn name(self) -> &'static str {
        match self {
            Self::SlicerInsert => "slicer_insert",
            Self::TimelineInsert => "timeline_insert",
            Self::SlicerReplace => "slicer_replace",
            Self::TimelineReplace => "timeline_replace",
        }
    }

    const fn is_slicer(self) -> bool {
        matches!(self, Self::SlicerInsert | Self::SlicerReplace)
    }

    const fn is_replace(self) -> bool {
        matches!(self, Self::SlicerReplace | Self::TimelineReplace)
    }
}

#[derive(Debug, Serialize)]
struct BinaryIdentity {
    path: String,
    bytes: usize,
    sha256: String,
}

#[derive(Debug, Serialize)]
struct ReadReport {
    operation: &'static str,
    selected_part: &'static str,
    elapsed_ns: Vec<u128>,
    allocation_samples: Vec<AllocationSample>,
    decode_counters_after: Option<(u64, u64)>,
}

#[derive(Debug, Serialize)]
struct OperationReport {
    operation: &'static str,
    base_bytes: usize,
    base_sha256: String,
    output_bytes: usize,
    output_sha256: String,
    elapsed_ns: Vec<u128>,
    allocation_samples: Vec<AllocationSample>,
    decode_counters_before: Option<(u64, u64)>,
    decode_counters_before_publication: Option<(u64, u64)>,
    decode_counters_after: Option<(u64, u64)>,
    native_identity: Vec<String>,
    untouched_worksheet_bytes: bool,
    replay_same_serialized_bytes: bool,
    inverse_restores_source_bytes: bool,
    stale_patch_refused: bool,
    no_op_preserves_source_bytes: bool,
}

#[derive(Debug, Serialize)]
struct CaseReport {
    name: &'static str,
    extra_parts: usize,
    extra_part_bytes: usize,
    cache_count: usize,
    open: ReadReport,
    first_read: ReadReport,
    operations: Vec<OperationReport>,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    api_variant: String,
    source_revision: String,
    corpus_generator: &'static str,
    lifecycle_scope: &'static str,
    iterations: usize,
    warmups: usize,
    binary: BinaryIdentity,
    cases: Vec<CaseReport>,
}

#[derive(Clone)]
enum PatchKind {
    Slicer(slicer::Patch),
    Timeline(timeline::Patch),
}

struct CommitData {
    package: Package,
    patch: PatchKind,
}

struct TimedCommit {
    data: CommitData,
    elapsed_ns: u128,
    allocation: AllocationSample,
    decode_counters_before_publication: Option<(u64, u64)>,
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

/// Open an owned archive with the revision's default OPC policy.
///
/// The 0bef/5726 source pair exposes the historical one-argument constructor,
/// whose implementation applies `SourceLimits::default()` internally.  The
/// current assembly keeps the same policy explicit at the required-limit
/// constructor.  This is the only versioned ingress adapter; the lifecycle
/// and ownership path around it stays identical.
// The current required-limit ingress is also the safe fallback for a
// no-default-features build.  The historical constructor is selected only
// when its explicit marker is present without the current marker.
#[cfg(any(feature = "current-xlsx-raw-opc", not(feature = "legacy-xlsx-raw-opc")))]
fn open_opc(bytes: Vec<u8>) -> ProbeResult<OpcPackage> {
    Ok(OpcPackage::from_vec(bytes, SourceLimits::default())?)
}

#[cfg(all(not(feature = "current-xlsx-raw-opc"), feature = "legacy-xlsx-raw-opc"))]
fn open_opc(bytes: Vec<u8>) -> ProbeResult<OpcPackage> {
    Ok(OpcPackage::from_vec(bytes)?)
}

fn open_package(bytes: &[u8]) -> ProbeResult<Package> {
    Ok(Package::from_opc(open_opc(bytes.to_vec())?)?)
}

fn cache_definition(index: usize, name: &str) -> ProbeResult<CacheDefinition> {
    Ok(CacheDefinition {
        name: name.to_owned(),
        uid: None,
        source_name: format!("Date{index}"),
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

fn slicer_xml(name: &str) -> String {
    format!(
        r#"<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="{name}" sourceName="State"><data><tabular pivotCacheId="5"><items count="2"><i x="1"/><i x="0" s="1"/></items></tabular></data></slicerCacheDefinition>"#
    )
}

fn raw_slicer_cache(index: usize, name: &str) -> ProbeResult<slicer::raw::Cache> {
    let limits = SourceLimits::default();
    Ok(slicer::raw::Cache {
        relationship_id: format!("rIdRawProbe{index}"),
        part_name: format!("/xl/slicerCaches/raw-probe-{index}.xml"),
        definition: slicer::raw::parse_cache(slicer_xml(name).as_bytes(), &limits)?,
    })
}

fn raw_timeline_cache(index: usize, name: &str) -> ProbeResult<timeline::raw::Cache> {
    let definition = cache_definition(index, name)?;
    let xml = timeline::codec::write_cache(&definition)?;
    let limits = SourceLimits::default();
    Ok(timeline::raw::Cache {
        relationship_id: format!("rIdRawProbe{index}"),
        part_name: format!("/xl/timelineCaches/raw-probe-{index}.xml"),
        definition: timeline::raw::parse_cache(&xml, &limits)?,
    })
}

fn raw_slicer_values(spec: CaseSpec) -> ProbeResult<Vec<slicer::raw::Cache>> {
    (0..spec.cache_count)
        .map(|index| raw_slicer_cache(index, &format!("RawProbeCache{index}")))
        .collect()
}

fn raw_timeline_values(spec: CaseSpec) -> ProbeResult<Vec<timeline::raw::Cache>> {
    (0..spec.cache_count)
        .map(|index| raw_timeline_cache(index, &format!("RawProbeCache{index}")))
        .collect()
}

fn commit_slicer(source: &Package, values: &[slicer::raw::Cache]) -> ProbeResult<CommitData> {
    let mut edit = source.edit_slicers()?;
    for value in values {
        edit.insert_cache(value.clone())?;
    }
    let commit = edit.commit()?;
    let patch = commit.patch().clone();
    Ok(CommitData {
        package: commit.into_package(),
        patch: PatchKind::Slicer(patch),
    })
}

fn commit_timeline(source: &Package, values: &[timeline::raw::Cache]) -> ProbeResult<CommitData> {
    let mut edit = source.edit_timelines()?;
    edit.insert_caches(values)?;
    let commit = edit.commit()?;
    let patch = commit.patch().clone();
    Ok(CommitData {
        package: commit.into_package(),
        patch: PatchKind::Timeline(patch),
    })
}

fn commit_slicer_replacement(source: &Package) -> ProbeResult<CommitData> {
    let mut edit = source.edit_slicers()?;
    let xml = slicer_xml(EXISTING_NAME).replace("sourceName=\"State\"", "sourceName=\"Changed\"");
    let value = slicer::raw::parse_cache(xml.as_bytes(), &SourceLimits::default())?;
    if !edit.replace_cache(EXISTING_NAME, value)? {
        return Err("slicer replacement did not find its native source name".into());
    }
    let commit = edit.commit()?;
    let patch = commit.patch().clone();
    Ok(CommitData {
        package: commit.into_package(),
        patch: PatchKind::Slicer(patch),
    })
}

fn commit_timeline_replacement(source: &Package) -> ProbeResult<CommitData> {
    let mut edit = source.edit_timelines()?;
    if !edit.replace_cache(EXISTING_NAME, cache_definition(99, EXISTING_NAME)?)? {
        return Err("timeline replacement did not find its native source name".into());
    }
    let commit = edit.commit()?;
    let patch = commit.patch().clone();
    Ok(CommitData {
        package: commit.into_package(),
        patch: PatchKind::Timeline(patch),
    })
}

fn apply_operation(
    source: &Package,
    operation: Operation,
    spec: CaseSpec,
) -> ProbeResult<CommitData> {
    match operation {
        Operation::SlicerInsert => commit_slicer(source, &raw_slicer_values(spec)?),
        Operation::TimelineInsert => commit_timeline(source, &raw_timeline_values(spec)?),
        Operation::SlicerReplace => commit_slicer_replacement(source),
        Operation::TimelineReplace => commit_timeline_replacement(source),
    }
}

fn build_source(spec: CaseSpec, operation: Operation) -> ProbeResult<Vec<u8>> {
    let source = build_package(spec)?;
    let source = if operation.is_replace() {
        if operation.is_slicer() {
            commit_slicer(
                &source,
                &[slicer::raw::Cache {
                    relationship_id: "rIdExistingRawProbe".into(),
                    part_name: "/xl/slicerCaches/existing-raw-probe.xml".into(),
                    definition: slicer::raw::parse_cache(
                        slicer_xml(EXISTING_NAME).as_bytes(),
                        &SourceLimits::default(),
                    )?,
                }],
            )?
            .package
        } else {
            commit_timeline(
                &source,
                &[timeline::raw::Cache {
                    relationship_id: "rIdExistingRawProbe".into(),
                    part_name: "/xl/timelineCaches/existing-raw-probe.xml".into(),
                    definition: timeline::raw::parse_cache(
                        &timeline::codec::write_cache(&cache_definition(0, EXISTING_NAME)?)?,
                        &SourceLimits::default(),
                    )?,
                }],
            )?
            .package
        }
    } else {
        source
    };
    Ok(source.to_bytes()?)
}

fn counters(package: &Package) -> Option<(u64, u64)> {
    package.clone().into_plain_opc().deferred_decode_counters()
}

fn part_bytes(package: &Package, path: &str) -> ProbeResult<Vec<u8>> {
    let opc = package.clone().into_plain_opc();
    Ok(opc.get_part(&PackURI::new(path)?)?.blob().to_vec())
}

fn untouched_large_parts(
    source: &Package,
    candidate: &Package,
    spec: CaseSpec,
) -> ProbeResult<bool> {
    for index in 0..spec.extra_parts {
        let path = format!("/xl/rawProbe/part-{index:03}.bin");
        let source_part = part_bytes(source, &path)?;
        let candidate_part = part_bytes(candidate, &path)?;
        if source_part != candidate_part {
            return Ok(false);
        }
    }
    Ok(true)
}

fn native_identity(
    package: &Package,
    operation: Operation,
    count: usize,
) -> ProbeResult<Vec<String>> {
    let opc = package.clone().into_plain_opc();
    let workbook = opc.get_part(&PackURI::new(WORKBOOK)?)?;
    let mut result = Vec::with_capacity(count);
    for index in 0..count {
        let id = if operation.is_replace() {
            "rIdExistingRawProbe".to_owned()
        } else {
            format!("rIdRawProbe{index}")
        };
        let (expected_path, expected_type) = if operation.is_slicer() {
            (
                if operation.is_replace() {
                    "/xl/slicerCaches/existing-raw-probe.xml".to_owned()
                } else {
                    format!("/xl/slicerCaches/raw-probe-{index}.xml")
                },
                slicer::CACHE_CONTENT_TYPE,
            )
        } else {
            (
                if operation.is_replace() {
                    "/xl/timelineCaches/existing-raw-probe.xml".to_owned()
                } else {
                    format!("/xl/timelineCaches/raw-probe-{index}.xml")
                },
                timeline::CACHE_CONTENT_TYPE,
            )
        };
        let relationship = workbook
            .rels()
            .get(&id)
            .ok_or_else(|| format!("missing native relationship {id}"))?;
        let target = relationship.target_partname()?.to_string();
        if target != expected_path {
            return Err(format!(
                "native relationship {id} targets {target}, expected {expected_path}"
            )
            .into());
        }
        let target_partname = relationship.target_partname()?;
        let part = opc.get_part(&target_partname)?;
        if part.content_type() != expected_type {
            return Err(format!(
                "native relationship {id} has content type {}, expected {expected_type}",
                part.content_type()
            )
            .into());
        }
        result.push(format!("{id}:{target}:{}", part.content_type()));
    }
    Ok(result)
}

fn inverse_patch(patch: &PatchKind) -> PatchKind {
    match patch {
        PatchKind::Slicer(patch) => PatchKind::Slicer(patch.inverse()),
        PatchKind::Timeline(patch) => PatchKind::Timeline(patch.inverse()),
    }
}

fn replay_patch(package: &Package, patch: &PatchKind) -> ProbeResult<Package> {
    Ok(match patch {
        PatchKind::Slicer(patch) => package.apply_slicers(patch)?.into_package(),
        PatchKind::Timeline(patch) => package.apply_timelines(patch)?.into_package(),
    })
}

fn no_op(source: &Package, operation: Operation, source_bytes: &[u8]) -> ProbeResult<bool> {
    if operation.is_slicer() {
        let commit = source.edit_slicers()?.commit()?;
        Ok(commit.patch().is_noop()
            && !commit.diagnostics().changed()
            && commit.package().to_bytes()? == source_bytes)
    } else {
        let commit = source.edit_timelines()?.commit()?;
        Ok(commit.patch().is_noop()
            && !commit.diagnostics().changed()
            && commit.package().to_bytes()? == source_bytes)
    }
}

fn verify_commit(
    source: &Package,
    source_bytes: &[u8],
    operation: Operation,
    spec: CaseSpec,
    data: &CommitData,
) -> ProbeResult<(
    Vec<u8>,
    Vec<String>,
    bool,
    bool,
    bool,
    bool,
    Option<(u64, u64)>,
)> {
    let before_publication = counters(&data.package);
    let output_bytes = data.package.to_bytes()?;
    let replay_bytes = replay_patch(source, &data.patch)?.to_bytes()?;
    let inverse_bytes = replay_patch(&data.package, &inverse_patch(&data.patch))?.to_bytes()?;
    let stale_refused = match replay_patch(&data.package, &data.patch) {
        Ok(_) => false,
        Err(error) => matches!(
            error.downcast_ref::<litchi_xlsx::Error>(),
            Some(litchi_xlsx::Error::PatchConflict { .. })
        ),
    };
    let count = if operation.is_replace() {
        1
    } else {
        spec.cache_count
    };
    let identities = native_identity(&data.package, operation, count)?;
    let untouched = part_bytes(source, WORKSHEET)? == part_bytes(&data.package, WORKSHEET)?;
    let reopened = open_package(&output_bytes)?;
    if reopened.to_bytes()? != output_bytes {
        return Err("reopened XLSX output changed serialized bytes".into());
    }
    if native_identity(&reopened, operation, count)? != identities {
        return Err("reopened XLSX output changed native relationship identity".into());
    }
    if !untouched_large_parts(source, &data.package, spec)?
        || !untouched_large_parts(source, &reopened, spec)?
    {
        return Err("a large untouched XLSX part changed across commit or reopen".into());
    }
    let noop = no_op(source, operation, source_bytes)?;
    if replay_bytes != output_bytes {
        return Err("replayed XLSX patch did not reproduce serialized output".into());
    }
    if inverse_bytes != source_bytes {
        return Err("inverse XLSX patch did not restore source bytes".into());
    }
    if !stale_refused {
        return Err("replaying an applied XLSX patch did not return PatchConflict".into());
    }
    if !noop || !untouched {
        return Err("XLSX no-op or untouched worksheet fidelity check failed".into());
    }
    Ok((
        output_bytes,
        identities,
        true,
        true,
        true,
        true,
        before_publication,
    ))
}

fn timed_commit(
    source_bytes: &[u8],
    operation: Operation,
    spec: CaseSpec,
) -> ProbeResult<TimedCommit> {
    let source = open_package(source_bytes)?;
    let region = allocation_metrics::begin();
    let started = Instant::now();
    let data = apply_operation(&source, operation, spec)?;
    let elapsed_ns = started.elapsed().as_nanos();
    let allocation = region
        .finish()
        .unwrap_or_else(allocation_metrics::unavailable_sample);
    let decode_counters_before_publication = counters(&data.package);
    Ok(TimedCommit {
        data,
        elapsed_ns,
        allocation,
        decode_counters_before_publication,
    })
}

fn open_report(source_bytes: &[u8], iterations: usize, warmups: usize) -> ProbeResult<ReadReport> {
    for _ in 0..warmups {
        let _ = open_package(source_bytes)?;
    }
    let mut elapsed_ns = Vec::with_capacity(iterations);
    let mut allocation_samples = Vec::with_capacity(iterations);
    let mut decode_counters_after = None;
    for _ in 0..iterations {
        let owned = source_bytes.to_vec();
        let region = allocation_metrics::begin();
        let started = Instant::now();
        let package = Package::from_opc(open_opc(owned)?)?;
        elapsed_ns.push(started.elapsed().as_nanos());
        allocation_samples.push(
            region
                .finish()
                .unwrap_or_else(allocation_metrics::unavailable_sample),
        );
        decode_counters_after = counters(&package);
    }
    Ok(ReadReport {
        operation: "owned_open",
        selected_part: WORKBOOK,
        elapsed_ns,
        allocation_samples,
        decode_counters_after,
    })
}

fn first_read_report(
    source_bytes: &[u8],
    iterations: usize,
    warmups: usize,
) -> ProbeResult<ReadReport> {
    let target = PackURI::new(WORKBOOK)?;
    for _ in 0..warmups {
        let package = open_opc(source_bytes.to_vec())?;
        let _ = package.get_part(&target)?.blob().len();
    }
    let mut elapsed_ns = Vec::with_capacity(iterations);
    let mut allocation_samples = Vec::with_capacity(iterations);
    let mut decode_counters_after = None;
    for _ in 0..iterations {
        let package = open_opc(source_bytes.to_vec())?;
        let region = allocation_metrics::begin();
        let started = Instant::now();
        let _ = package.get_part(&target)?.blob().len();
        elapsed_ns.push(started.elapsed().as_nanos());
        allocation_samples.push(
            region
                .finish()
                .unwrap_or_else(allocation_metrics::unavailable_sample),
        );
        decode_counters_after = package.deferred_decode_counters();
    }
    Ok(ReadReport {
        operation: "first_selected_read",
        selected_part: WORKBOOK,
        elapsed_ns,
        allocation_samples,
        decode_counters_after,
    })
}

fn operation_report(
    source_bytes: &[u8],
    operation: Operation,
    spec: CaseSpec,
    iterations: usize,
    warmups: usize,
) -> ProbeResult<OperationReport> {
    let source = open_package(source_bytes)?;
    let before = counters(&source);
    let expected = apply_operation(&source, operation, spec)?;
    let (
        output_bytes,
        identities,
        replay_ok,
        inverse_ok,
        stale_refused,
        no_op_ok,
        before_publication,
    ) = verify_commit(&source, source_bytes, operation, spec, &expected)?;
    for _ in 0..warmups {
        drop(timed_commit(source_bytes, operation, spec)?);
    }
    let mut elapsed_ns = Vec::with_capacity(iterations);
    let mut allocation_samples = Vec::with_capacity(iterations);
    let mut after = None;
    for _ in 0..iterations {
        let timed = timed_commit(source_bytes, operation, spec)?;
        let candidate_bytes = timed.data.package.to_bytes()?;
        if candidate_bytes != output_bytes {
            return Err(format!("non-deterministic output for {}", operation.name()).into());
        }
        after = timed.decode_counters_before_publication;
        elapsed_ns.push(timed.elapsed_ns);
        allocation_samples.push(timed.allocation);
    }
    Ok(OperationReport {
        operation: operation.name(),
        base_bytes: source_bytes.len(),
        base_sha256: sha256_hex(source_bytes),
        output_bytes: output_bytes.len(),
        output_sha256: sha256_hex(&output_bytes),
        elapsed_ns,
        allocation_samples,
        decode_counters_before: before,
        decode_counters_before_publication: before_publication,
        decode_counters_after: after,
        native_identity: identities,
        untouched_worksheet_bytes: part_bytes(&source, WORKSHEET)?
            == part_bytes(&expected.package, WORKSHEET)?,
        replay_same_serialized_bytes: replay_ok,
        inverse_restores_source_bytes: inverse_ok,
        stale_patch_refused: stale_refused,
        no_op_preserves_source_bytes: no_op_ok,
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

/// Run the matched raw lifecycle probe. The source revision is required so a
/// report cannot be confused with a measurement from an unrecorded checkout.
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
    let case_reports = cases
        .into_iter()
        .map(|case| {
            let open_source = build_source(case, Operation::SlicerInsert)?;
            let open = open_report(&open_source, iterations, warmups)?;
            let first_read = first_read_report(&open_source, iterations, warmups)?;
            let operations = Operation::ALL
                .into_iter()
                .map(|operation| {
                    let source = build_source(case, operation)?;
                    operation_report(&source, operation, case, iterations, warmups)
                })
                .collect::<ProbeResult<Vec<_>>>()?;
            Ok(CaseReport {
                name: case.name,
                extra_parts: case.extra_parts,
                extra_part_bytes: case.extra_part_bytes,
                cache_count: case.cache_count,
                open,
                first_read,
                operations,
            })
        })
        .collect::<ProbeResult<Vec<_>>>()?;
    let report = Report {
        schema: SCHEMA,
        api_variant,
        source_revision,
        corpus_generator: CORPUS_GENERATOR,
        lifecycle_scope: "owned open, first selected read, raw slicer/timeline insertion, typed replacement, commit, replay, inverse, stale patch, no-op, native identity, untouched worksheet bytes and deferred counters",
        iterations,
        warmups,
        binary: binary_identity()?,
        cases: case_reports,
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
