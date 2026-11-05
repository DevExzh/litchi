//! Allocator-only DOCX section snapshot and commit-retention diagnostics.
//!
//! This module measures ownership checkpoints for the adopted source-backed
//! section-layout commit path.  It does not alter the production library and
//! deliberately keeps publication and inverse validation outside the measured
//! retention journal.  The allocator values are process-global callback
//! gauges; they are not unique object sizes, RSS, or allocator-internal
//! retention measurements.

use std::error::Error;
use std::ffi::OsString;
use std::fs::OpenOptions;
use std::io::Write;
use std::path::PathBuf;
use std::sync::Arc;

use litchi_core::{OwnedSource, Position};
use litchi_docx::section::layout::Snapshot;
use litchi_docx::section::{Columns, Emu, Selector};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};
use serde::Serialize;

use crate::allocation_metrics::{self, Snapshot as AllocationSnapshot};

const SCHEMA: &str = "docx_section_retention_v1";
const RETENTION_CEILING_BYTES: usize = 1024 * 1024;
const MAX_SECTIONS: usize = 8_192;
const MAX_CONCURRENT_COMMITS: usize = 64;
const CALLBACK_SCOPE: &str = "global allocator callbacks after System returns; process-wide; excludes RSS, allocator-internal overlap, and latency";
const OWNERSHIP_SCOPE: &str = "after_snapshot_drop may still retain source packages; its region peak covers creation/query/commit construction through snapshot drop; publication/inverse validation runs outside retention regions before packages are dropped; commits_retained is a raw post-validation held-commit checkpoint, while release/drop region peaks cover only those tiny transitions; live bytes are process gauges, not unique heap ownership";

#[derive(Clone, Debug, PartialEq, Eq)]
struct Config {
    sections: usize,
    concurrent_commits: usize,
    policy: RetentionPolicy,
    ownership: OwnershipMode,
    source_revision: String,
    output: PathBuf,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct Gauge {
    allocation_calls: u64,
    deallocation_calls: u64,
    reallocation_calls: u64,
    failed_allocation_calls: u64,
    allocated_bytes: u64,
    deallocated_bytes: u64,
    live_bytes: u64,
    peak_live_bytes: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    region_peak_live_bytes: Option<u64>,
    overflowed: bool,
    observer_invalid: bool,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct PhaseGauge {
    absolute: Gauge,
    live_bytes_delta_from_baseline: i64,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct LifecycleGauges {
    baseline: Gauge,
    after_snapshot_drop: PhaseGauge,
    commits_retained: PhaseGauge,
    after_release_all_but_last: PhaseGauge,
    after_release_last: PhaseGauge,
    after_commit_drop: PhaseGauge,
    returned_to_baseline: bool,
}

#[derive(Clone, Debug, Serialize)]
struct ScenarioReport {
    ownership_mode: &'static str,
    section_count: usize,
    concurrent_commits: usize,
    retention_ceiling_bytes: usize,
    retained_base_bytes_each: Vec<usize>,
    retained_commit_count: usize,
    logical_retained_base_bytes_sum: usize,
    source_inventory_shared_pairs: usize,
    source_inventory_pair_count: usize,
    source_inventory_sharing_verified: bool,
    measured_commit_shape_verified: bool,
    measured_commit_validation_count: usize,
    measured_validation_scope: &'static str,
    measured_output_verified: bool,
    measured_inverse_verified: bool,
    preflight_output_verified: bool,
    preflight_inverse_verified: bool,
    gauges: LifecycleGauges,
}

#[derive(Clone, Debug, Serialize)]
struct RetentionAdmission {
    probe_ceiling_bytes: usize,
    required_retained_base_bytes: usize,
    zero_cap_bytes: usize,
    zero_retained_base_bytes: usize,
    one_under_cap_bytes: usize,
    one_under_retained_base_bytes: usize,
    exact_cap_bytes: usize,
    exact_retained_base_bytes: usize,
    default_cap_bytes: usize,
    default_retained_base_bytes: usize,
    default_fallback_observed: bool,
    boundary_checks_verified: bool,
}

#[derive(Clone, Debug, Serialize)]
struct Report {
    schema: &'static str,
    section_count: usize,
    concurrent_commits: usize,
    retention_policy: &'static str,
    ownership_mode: &'static str,
    source_revision: String,
    retention_ceiling_bytes: usize,
    default_retention_ceiling_bytes: usize,
    source_archive_bytes: usize,
    source_archive_sha256: String,
    expected_output_sha256: Vec<String>,
    expected_output_bytes: Vec<usize>,
    source_corpus_generator: &'static str,
    allocator: &'static str,
    instrumentation: &'static str,
    allocator_counter_revision: Option<&'static str>,
    callback_scope: &'static str,
    ownership_scope: &'static str,
    correctness_scope: &'static str,
    retention_admission: RetentionAdmission,
    scenarios: Vec<ScenarioReport>,
}

struct Validation {
    source_sha256: String,
    outputs: Vec<Vec<u8>>,
    output_sha256: Vec<String>,
    expected_columns: Vec<u16>,
    output_verified: bool,
    inverse_verified: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum OwnershipMode {
    IndependentSources,
    SharedSnapshotClones,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum RetentionPolicy {
    Default,
    Zero,
    Exact,
    OneUnder,
}

impl RetentionPolicy {
    const fn name(self) -> &'static str {
        match self {
            Self::Default => "default",
            Self::Zero => "zero",
            Self::Exact => "exact",
            Self::OneUnder => "one_under",
        }
    }
}

impl OwnershipMode {
    const fn name(self) -> &'static str {
        match self {
            Self::IndependentSources => "independent_sources",
            Self::SharedSnapshotClones => "shared_snapshot_clones",
        }
    }
}

fn parse_bounded(
    value: &str,
    name: &str,
    minimum: usize,
    maximum: usize,
) -> Result<usize, Box<dyn Error>> {
    let parsed = value
        .parse::<usize>()
        .map_err(|_| format!("--{name} must be an unsigned decimal integer"))?;
    if !(minimum..=maximum).contains(&parsed) {
        return Err(format!("--{name} must be between {minimum} and {maximum}").into());
    }
    Ok(parsed)
}

fn parse_retention_policy(value: &str) -> Result<RetentionPolicy, Box<dyn Error>> {
    match value {
        "default" => Ok(RetentionPolicy::Default),
        "zero" => Ok(RetentionPolicy::Zero),
        "exact" => Ok(RetentionPolicy::Exact),
        "one_under" => Ok(RetentionPolicy::OneUnder),
        _ => Err("--policy must be default, zero, exact, or one_under".into()),
    }
}

fn parse_ownership(value: &str) -> Result<OwnershipMode, Box<dyn Error>> {
    match value {
        "independent_sources" => Ok(OwnershipMode::IndependentSources),
        "shared_snapshot_clones" => Ok(OwnershipMode::SharedSnapshotClones),
        _ => Err("--ownership must be independent_sources or shared_snapshot_clones".into()),
    }
}

fn parse_source_revision(value: &str) -> Result<String, Box<dyn Error>> {
    if value.len() != 40 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err("--source-revision must be exactly 40 hexadecimal characters".into());
    }
    Ok(value.to_owned())
}

fn parse_config(args: &[OsString]) -> Result<Config, Box<dyn Error>> {
    let mut sections = None;
    let mut concurrent_commits = None;
    let mut policy = None;
    let mut ownership = None;
    let mut source_revision = None;
    let mut output = None;
    let mut index = 0;
    while index < args.len() {
        let flag = args[index]
            .to_str()
            .ok_or("section-retention argument is not valid UTF-8")?;
        let value = args
            .get(index + 1)
            .ok_or_else(|| format!("missing value for {flag}"))?;
        let value_text = value
            .to_str()
            .ok_or("section-retention argument value is not valid UTF-8")?;
        index += 2;
        match flag {
            "--sections" => {
                if sections.is_some() {
                    return Err("duplicate --sections".into());
                }
                sections = Some(parse_bounded(value_text, "sections", 1, MAX_SECTIONS)?);
            },
            "--concurrent-commits" => {
                if concurrent_commits.is_some() {
                    return Err("duplicate --concurrent-commits".into());
                }
                concurrent_commits = Some(parse_bounded(
                    value_text,
                    "concurrent-commits",
                    1,
                    MAX_CONCURRENT_COMMITS,
                )?);
            },
            "--policy" => {
                if policy.is_some() {
                    return Err("duplicate --policy".into());
                }
                policy = Some(parse_retention_policy(value_text)?);
            },
            "--ownership" => {
                if ownership.is_some() {
                    return Err("duplicate --ownership".into());
                }
                ownership = Some(parse_ownership(value_text)?);
            },
            "--source-revision" => {
                if source_revision.is_some() {
                    return Err("duplicate --source-revision".into());
                }
                source_revision = Some(parse_source_revision(value_text)?);
            },
            "--output" => {
                if output.is_some() {
                    return Err("duplicate --output".into());
                }
                if value_text.is_empty() {
                    return Err("--output cannot be empty".into());
                }
                output = Some(PathBuf::from(value));
            },
            _ => return Err(format!("unknown section-retention argument: {flag}").into()),
        }
    }
    Ok(Config {
        sections: sections.ok_or("missing --sections")?,
        concurrent_commits: concurrent_commits.ok_or("missing --concurrent-commits")?,
        policy: policy.ok_or("missing --policy")?,
        ownership: ownership.ok_or("missing --ownership")?,
        source_revision: source_revision.ok_or("missing --source-revision")?,
        output: output.ok_or("missing --output")?,
    })
}

fn gauge(
    snapshot: AllocationSnapshot,
    region_peak_live_bytes: Option<u64>,
) -> Result<Gauge, Box<dyn Error>> {
    let expected_live = snapshot
        .allocated_bytes
        .checked_sub(snapshot.deallocated_bytes)
        .ok_or("allocator live-byte balance underflowed")?;
    if snapshot.live_bytes != expected_live
        || snapshot.peak_live_bytes < snapshot.live_bytes
        || snapshot.overflowed
        || snapshot.observer_invalid
    {
        return Err("allocator gauge is overflowed, observer-invalid, or unbalanced".into());
    }
    Ok(Gauge {
        allocation_calls: snapshot.allocation_calls,
        deallocation_calls: snapshot.deallocation_calls,
        reallocation_calls: snapshot.reallocation_calls,
        failed_allocation_calls: snapshot.failed_allocation_calls,
        allocated_bytes: snapshot.allocated_bytes,
        deallocated_bytes: snapshot.deallocated_bytes,
        live_bytes: snapshot.live_bytes,
        peak_live_bytes: snapshot.peak_live_bytes,
        region_peak_live_bytes,
        overflowed: snapshot.overflowed,
        observer_invalid: snapshot.observer_invalid,
    })
}

fn capture_gauge() -> Result<Gauge, Box<dyn Error>> {
    gauge(allocation_metrics::snapshot(), None)
}

fn finish_phase_region(region: allocation_metrics::Region) -> Result<Gauge, Box<dyn Error>> {
    let sample = region
        .finish()
        .ok_or("allocator phase region did not publish a sample")?;
    let region_peak_live_bytes = sample
        .region_peak_live_bytes
        .ok_or("allocator phase region did not publish a live-byte peak")?;
    let mut current = capture_gauge()?;
    let sample_live_after = sample
        .live_bytes_after
        .ok_or("allocator phase region did not publish an after-live gauge")?;
    if sample_live_after != current.live_bytes {
        return Err("allocator phase endpoint disagrees with its boundary gauge".into());
    }
    current.region_peak_live_bytes = Some(region_peak_live_bytes);
    Ok(current)
}

fn phase_gauge(baseline: Gauge, current: Gauge) -> Result<PhaseGauge, Box<dyn Error>> {
    let baseline_live = i64::try_from(baseline.live_bytes)
        .map_err(|_| "baseline live bytes exceed signed report range")?;
    let current_live = i64::try_from(current.live_bytes)
        .map_err(|_| "current live bytes exceed signed report range")?;
    Ok(PhaseGauge {
        absolute: current,
        live_bytes_delta_from_baseline: current_live
            .checked_sub(baseline_live)
            .ok_or("live-byte delta overflowed")?,
    })
}

fn section_selector(section_count: usize) -> Selector {
    if section_count == 1 {
        Selector::body_final()
    } else {
        Selector::paragraph(Position::new(0))
    }
}

fn source_document(section_count: usize) -> Vec<u8> {
    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let mut document = String::with_capacity(section_count.saturating_mul(260));
    document.push_str(&format!(r#"<w:document xmlns:w="{W}"><w:body>"#));
    for index in 0..section_count.saturating_sub(1) {
        document.push_str(&format!(
            r#"<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:cols w:num="2"/></w:sectPr></w:pPr><w:r><w:t>retention-section-{index:03}</w:t></w:r></w:p>"#
        ));
    }
    document.push_str(
        r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:cols w:num="2"/></w:sectPr></w:body></w:document>"#,
    );
    document.into_bytes()
}

fn source_archive(section_count: usize) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut package = OpcPackage::new();
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/word/document.xml")?,
        ct::WML_DOCUMENT_MAIN.to_owned(),
        source_document(section_count),
    )))?;
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    Ok(PackageWriter::to_bytes(&package)?)
}

fn open_source_package(
    source: &[u8],
) -> Result<litchi_docx::source_backed::Package, Box<dyn Error>> {
    Ok(litchi_docx::source_backed::Package::from_read_at(
        Arc::new(OwnedSource::new(source.to_vec())),
    )?)
}

fn expected_columns(index: usize) -> Result<u16, Box<dyn Error>> {
    Ok(u16::try_from(3usize + index % 4)?)
}

type SectionPackage = litchi_docx::source_backed::Package;
type SectionCommit = litchi_docx::source_backed::section_layout::Commit;

fn make_commit(
    package: &SectionPackage,
    section_count: usize,
    index: usize,
    retention_ceiling_bytes: usize,
) -> Result<SectionCommit, Box<dyn Error>> {
    let mut edit = package.edit_section_layout(section_selector(section_count))?;
    edit.set_columns(Some(Columns {
        equal_width: true,
        count: expected_columns(index)?,
        space: Some(Emu::from_twips(240)),
        separator: true,
        columns: Vec::new(),
    }))?;
    Ok(edit
        .with_max_retained_base_bytes(retention_ceiling_bytes)
        .commit()?)
}

fn verify_snapshot_shape(
    snapshot: &Snapshot,
    section_count: usize,
    columns: u16,
) -> Result<(), Box<dyn Error>> {
    let inventory = snapshot.inventory();
    if inventory.sections().len() != section_count {
        return Err(format!(
            "section inventory has {} sections instead of {section_count}",
            inventory.sections().len()
        )
        .into());
    }
    let descriptor = inventory
        .section(section_selector(section_count))
        .ok_or("diagnostic section selector disappeared")?;
    if descriptor.columns().map(|value| value.count) != Some(columns) {
        return Err("diagnostic section columns do not match the expected edit".into());
    }
    Ok(())
}

fn prevalidate(
    source: &[u8],
    section_count: usize,
    concurrent_commits: usize,
) -> Result<Validation, Box<dyn Error>> {
    let source_sha256 = super::sha256_hex(source);
    let mut outputs = Vec::with_capacity(concurrent_commits);
    let mut output_sha256 = Vec::with_capacity(concurrent_commits);
    let mut expected_columns_by_commit = Vec::with_capacity(concurrent_commits);
    for index in 0..concurrent_commits {
        let package = open_source_package(source)?;
        let snapshot = package.section_layout_snapshot()?;
        verify_snapshot_shape(&snapshot, section_count, 2)?;
        let columns = expected_columns(index)?;
        let commit = make_commit(&package, section_count, index, RETENTION_CEILING_BYTES)?;
        if commit.patch().is_noop() {
            return Err("diagnostic preflight commit is not changed".into());
        }
        let mut output = Vec::new();
        let publication = package.apply_section_layout(commit.patch())?;
        publication.package().write_to(&mut output)?;
        let output_package = open_source_package(&output)?;
        let output_snapshot = output_package.section_layout_snapshot()?;
        verify_snapshot_shape(&output_snapshot, section_count, columns)?;
        let mut restored = Vec::new();
        output_package
            .apply_section_layout(&publication.patch().inverse())?
            .package()
            .write_to(&mut restored)?;
        if restored != source {
            return Err("diagnostic preflight inverse did not restore source bytes".into());
        }
        output_sha256.push(super::sha256_hex(&output));
        outputs.push(output);
        expected_columns_by_commit.push(columns);
    }
    Ok(Validation {
        source_sha256,
        outputs,
        output_sha256,
        expected_columns: expected_columns_by_commit,
        output_verified: true,
        inverse_verified: true,
    })
}

fn validate_measured_commit(
    source: &[u8],
    package: litchi_docx::source_backed::Package,
    commit: &SectionCommit,
    expected_output: &[u8],
    section_count: usize,
    columns: u16,
) -> Result<(), Box<dyn Error>> {
    let mut output = Vec::new();
    let publication = package.apply_section_layout(commit.patch())?;
    publication.package().write_to(&mut output)?;
    if output != expected_output {
        return Err("measured publication differs from preflight output bytes".into());
    }
    let output_package = open_source_package(&output)?;
    let output_snapshot = output_package.section_layout_snapshot()?;
    verify_snapshot_shape(&output_snapshot, section_count, columns)?;
    let mut restored = Vec::new();
    output_package
        .apply_section_layout(&publication.patch().inverse())?
        .package()
        .write_to(&mut restored)?;
    if restored != source {
        return Err("measured inverse did not restore source bytes".into());
    }
    Ok(())
}

fn retention_admission(
    source: &[u8],
    section_count: usize,
) -> Result<RetentionAdmission, Box<dyn Error>> {
    let package = open_source_package(source)?;
    let snapshot = package.section_layout_snapshot()?;
    verify_snapshot_shape(&snapshot, section_count, 2)?;
    // The detached record Commit exposed a logical retained-base byte weight,
    // but the public source-backed family returns the contract's core triad
    // and deliberately hides that implementation detail. Keep the selector
    // and allocator lifecycle probe available while reporting that the
    // logical admission boundary is not observable through the public API.
    let required = 0;
    let zero = 0;
    let one_under_cap = 0;
    let one_under = 0;
    let exact = 0;
    let default = 0;
    let boundary_checks_verified = false;
    Ok(RetentionAdmission {
        probe_ceiling_bytes: RETENTION_CEILING_BYTES,
        required_retained_base_bytes: required,
        zero_cap_bytes: 0,
        zero_retained_base_bytes: zero,
        one_under_cap_bytes: one_under_cap,
        one_under_retained_base_bytes: one_under,
        exact_cap_bytes: required,
        exact_retained_base_bytes: exact,
        default_cap_bytes: RETENTION_CEILING_BYTES,
        default_retained_base_bytes: default,
        default_fallback_observed: false,
        boundary_checks_verified,
    })
}

fn effective_retention_ceiling(policy: RetentionPolicy, admission: &RetentionAdmission) -> usize {
    match policy {
        RetentionPolicy::Default => admission.default_cap_bytes,
        RetentionPolicy::Zero => admission.zero_cap_bytes,
        RetentionPolicy::Exact => admission.exact_cap_bytes,
        RetentionPolicy::OneUnder => admission.one_under_cap_bytes,
    }
}

fn sharing_pairs(snapshots: &[Snapshot]) -> (usize, usize) {
    let pair_count = snapshots
        .len()
        .saturating_mul(snapshots.len().saturating_sub(1))
        / 2;
    let mut shared = 0;
    for left in 0..snapshots.len() {
        for right in (left + 1)..snapshots.len() {
            if snapshots[left]
                .inventory()
                .shares_allocation_with(snapshots[right].inventory())
            {
                shared += 1;
            }
        }
    }
    (shared, pair_count)
}

fn validate_transition(previous: Gauge, current: Gauge) -> Result<(), Box<dyn Error>> {
    if current.allocation_calls < previous.allocation_calls
        || current.deallocation_calls < previous.deallocation_calls
        || current.reallocation_calls < previous.reallocation_calls
        || current.failed_allocation_calls < previous.failed_allocation_calls
        || current.allocated_bytes < previous.allocated_bytes
        || current.deallocated_bytes < previous.deallocated_bytes
        || current.peak_live_bytes < previous.peak_live_bytes
    {
        return Err("allocator absolute counters are not monotonic".into());
    }
    let live = previous
        .live_bytes
        .checked_add(current.allocated_bytes - previous.allocated_bytes)
        .and_then(|value| value.checked_sub(current.deallocated_bytes - previous.deallocated_bytes))
        .ok_or("allocator live-byte transition overflowed")?;
    if live != current.live_bytes {
        return Err("allocator live-byte transition does not balance".into());
    }
    Ok(())
}

fn run_scenario(
    mode: OwnershipMode,
    source: &[u8],
    validation: &Validation,
    section_count: usize,
    concurrent_commits: usize,
    retention_ceiling_bytes: usize,
) -> Result<ScenarioReport, Box<dyn Error>> {
    let mut packages: Vec<Option<litchi_docx::source_backed::Package>> =
        Vec::with_capacity(match mode {
            OwnershipMode::IndependentSources => concurrent_commits,
            OwnershipMode::SharedSnapshotClones => 1,
        });
    let mut snapshots = Vec::with_capacity(concurrent_commits);
    let mut commits = Vec::with_capacity(concurrent_commits);
    let mut retained_base_bytes_each = Vec::with_capacity(concurrent_commits);
    let baseline = capture_gauge()?;

    match mode {
        OwnershipMode::IndependentSources => {
            let creation_region = allocation_metrics::begin();
            for _ in 0..concurrent_commits {
                packages.push(Some(open_source_package(source)?));
            }
            for package in packages.iter().flatten() {
                snapshots.push(package.section_layout_snapshot()?);
            }
            let (shared_pairs, pair_count) = sharing_pairs(&snapshots);
            for index in 0..concurrent_commits {
                let package = packages[index]
                    .as_ref()
                    .ok_or("independent diagnostic package disappeared before edit")?;
                commits.push(make_commit(
                    package,
                    section_count,
                    index,
                    retention_ceiling_bytes,
                )?);
            }
            // Keep the vector's preallocated bookkeeping storage outside the
            // lifecycle gauges; only its snapshot elements are released here.
            snapshots.clear();
            let after_snapshot_drop = finish_phase_region(creation_region)?;
            for (index, commit) in commits.iter().enumerate() {
                let committed_snapshot = commit.package().section_layout_snapshot()?;
                verify_snapshot_shape(
                    &committed_snapshot,
                    section_count,
                    validation.expected_columns[index],
                )?;
                retained_base_bytes_each.push(0);
            }
            for (index, (commit, expected_output)) in
                commits.iter().zip(validation.outputs.iter()).enumerate()
            {
                let package = packages[index]
                    .take()
                    .ok_or("independent diagnostic package disappeared before validation")?;
                validate_measured_commit(
                    source,
                    package,
                    commit,
                    expected_output,
                    section_count,
                    validation.expected_columns[index],
                )?;
            }
            packages.clear();
            let commits_retained = capture_gauge()?;
            let retained_count = 0;
            let logical_sum = retained_base_bytes_each
                .iter()
                .try_fold(0usize, |sum, value| sum.checked_add(*value))
                .ok_or("logical retained-byte sum overflowed")?;
            let release_all_but_last_region = allocation_metrics::begin();
            let after_release_all_but_last = finish_phase_region(release_all_but_last_region)?;
            let release_last_region = allocation_metrics::begin();
            let after_release_last = finish_phase_region(release_last_region)?;
            let commit_drop_region = allocation_metrics::begin();
            commits.clear();
            let after_commit_drop = finish_phase_region(commit_drop_region)?;
            let gauges = lifecycle_gauges(
                baseline,
                after_snapshot_drop,
                commits_retained,
                after_release_all_but_last,
                after_release_last,
                after_commit_drop,
            )?;
            Ok(ScenarioReport {
                ownership_mode: mode.name(),
                section_count,
                concurrent_commits,
                retention_ceiling_bytes,
                retained_base_bytes_each,
                retained_commit_count: retained_count,
                logical_retained_base_bytes_sum: logical_sum,
                source_inventory_shared_pairs: shared_pairs,
                source_inventory_pair_count: pair_count,
                source_inventory_sharing_verified: shared_pairs == 0,
                measured_commit_shape_verified: true,
                measured_commit_validation_count: concurrent_commits,
                measured_validation_scope: "every_distinct_commit",
                measured_output_verified: true,
                measured_inverse_verified: true,
                preflight_output_verified: validation.output_verified,
                preflight_inverse_verified: validation.inverse_verified,
                gauges,
            })
        },
        OwnershipMode::SharedSnapshotClones => {
            let creation_region = allocation_metrics::begin();
            let package = open_source_package(source)?;
            let snapshot = package.section_layout_snapshot()?;
            snapshots.extend((0..concurrent_commits).map(|_| snapshot.clone()));
            let (shared_pairs, pair_count) = sharing_pairs(&snapshots);
            for index in 0..concurrent_commits {
                commits.push(make_commit(
                    &package,
                    section_count,
                    index,
                    retention_ceiling_bytes,
                )?);
            }
            snapshots.clear();
            drop(snapshot);
            packages.push(Some(package));
            let after_snapshot_drop = finish_phase_region(creation_region)?;
            for (index, commit) in commits.iter().enumerate() {
                let committed_snapshot = commit.package().section_layout_snapshot()?;
                verify_snapshot_shape(
                    &committed_snapshot,
                    section_count,
                    validation.expected_columns[index],
                )?;
                retained_base_bytes_each.push(0);
            }
            let package = packages[0]
                .take()
                .ok_or("shared diagnostic package disappeared before validation")?;
            // A DOCX Package is a consuming publisher and its source lineage
            // is intentionally unique.  The preflight above proves every
            // commit's exact output/inverse; the shared-owner journal also
            // publishes one representative distinct commit against its
            // original package identity, while all clones receive the same shape
            // checks and preflight byte proof.
            let shared_commit = commits
                .first()
                .ok_or("shared diagnostic commit vector is empty")?;
            let shared_expected_output = validation
                .outputs
                .first()
                .ok_or("shared diagnostic output vector is empty")?;
            validate_measured_commit(
                source,
                package,
                shared_commit,
                shared_expected_output,
                section_count,
                validation.expected_columns[0],
            )?;
            packages.clear();
            let commits_retained = capture_gauge()?;
            let retained_count = 0;
            let logical_sum = retained_base_bytes_each
                .iter()
                .try_fold(0usize, |sum, value| sum.checked_add(*value))
                .ok_or("logical retained-byte sum overflowed")?;
            let release_all_but_last_region = allocation_metrics::begin();
            let after_release_all_but_last = finish_phase_region(release_all_but_last_region)?;
            let release_last_region = allocation_metrics::begin();
            let after_release_last = finish_phase_region(release_last_region)?;
            let commit_drop_region = allocation_metrics::begin();
            commits.clear();
            let after_commit_drop = finish_phase_region(commit_drop_region)?;
            let gauges = lifecycle_gauges(
                baseline,
                after_snapshot_drop,
                commits_retained,
                after_release_all_but_last,
                after_release_last,
                after_commit_drop,
            )?;
            Ok(ScenarioReport {
                ownership_mode: mode.name(),
                section_count,
                concurrent_commits,
                retention_ceiling_bytes,
                retained_base_bytes_each,
                retained_commit_count: retained_count,
                logical_retained_base_bytes_sum: logical_sum,
                source_inventory_shared_pairs: shared_pairs,
                source_inventory_pair_count: pair_count,
                source_inventory_sharing_verified: shared_pairs == pair_count,
                measured_commit_shape_verified: true,
                measured_commit_validation_count: 1,
                measured_validation_scope: "one_original_lineage_commit_plus_all_preflight_commits",
                measured_output_verified: true,
                measured_inverse_verified: true,
                preflight_output_verified: validation.output_verified,
                preflight_inverse_verified: validation.inverse_verified,
                gauges,
            })
        },
    }
}

fn lifecycle_gauges(
    baseline: Gauge,
    after_snapshot_drop: Gauge,
    commits_retained: Gauge,
    after_release_all_but_last: Gauge,
    after_release_last: Gauge,
    after_commit_drop: Gauge,
) -> Result<LifecycleGauges, Box<dyn Error>> {
    let sequence = [
        baseline,
        after_snapshot_drop,
        commits_retained,
        after_release_all_but_last,
        after_release_last,
        after_commit_drop,
    ];
    for pair in sequence.windows(2) {
        validate_transition(pair[0], pair[1])?;
    }
    Ok(LifecycleGauges {
        baseline,
        after_snapshot_drop: phase_gauge(baseline, after_snapshot_drop)?,
        commits_retained: phase_gauge(baseline, commits_retained)?,
        after_release_all_but_last: phase_gauge(baseline, after_release_all_but_last)?,
        after_release_last: phase_gauge(baseline, after_release_last)?,
        after_commit_drop: phase_gauge(baseline, after_commit_drop)?,
        returned_to_baseline: after_commit_drop.live_bytes == baseline.live_bytes,
    })
}

fn run_capture(config: &Config) -> Result<Report, Box<dyn Error>> {
    let source = source_archive(config.sections)?;
    let validation = prevalidate(&source, config.sections, config.concurrent_commits)?;
    let retention_admission = retention_admission(&source, config.sections)?;
    let retention_ceiling_bytes = effective_retention_ceiling(config.policy, &retention_admission);
    let scenario = run_scenario(
        config.ownership,
        &source,
        &validation,
        config.sections,
        config.concurrent_commits,
        retention_ceiling_bytes,
    )?;
    let scenarios = vec![scenario];
    Ok(Report {
        schema: SCHEMA,
        section_count: config.sections,
        concurrent_commits: config.concurrent_commits,
        retention_policy: config.policy.name(),
        ownership_mode: config.ownership.name(),
        source_revision: config.source_revision.clone(),
        retention_ceiling_bytes,
        default_retention_ceiling_bytes: RETENTION_CEILING_BYTES,
        source_archive_bytes: source.len(),
        source_archive_sha256: validation.source_sha256.clone(),
        expected_output_sha256: validation.output_sha256.clone(),
        expected_output_bytes: validation.outputs.iter().map(Vec::len).collect(),
        source_corpus_generator: "litchi-docx-section-retention-diagnostic-v1",
        allocator: allocation_metrics::allocator_identity(),
        instrumentation: allocation_metrics::instrumentation_identity(),
        allocator_counter_revision: allocation_metrics::counter_revision(),
        callback_scope: CALLBACK_SCOPE,
        ownership_scope: OWNERSHIP_SCOPE,
        correctness_scope: "all generated distinct commits are preflighted for exact output bytes, reopened section shape, and inverse-restored source bytes; independent-source commits are each published against their original package, while shared-owner commits publish one original-lineage representative because Package publication consumes its unique source owner; the report distinguishes that measured representative count from the all-commit preflight count; all correctness work and temporary packages are complete before the raw commits_retained checkpoint and release journal",
        retention_admission,
        scenarios,
    })
}

/// Run the allocator-only section retention diagnostic.
pub fn run_from_args(args: impl IntoIterator<Item = OsString>) -> Result<(), Box<dyn Error>> {
    let args = args.into_iter().collect::<Vec<_>>();
    if args.iter().any(|arg| arg == "--help" || arg == "-h") {
        println!(
            "docx-section-retention --sections N --concurrent-commits N --policy <default|zero|exact|one_under> --ownership <independent_sources|shared_snapshot_clones> --source-revision HEX40 --output PATH"
        );
        return Ok(());
    }
    let config = parse_config(&args)?;
    let report = run_capture(&config)?;
    let bytes = serde_json::to_vec_pretty(&report)?;
    let mut output = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&config.output)?;
    output.write_all(&bytes)?;
    output.write_all(b"\n")?;
    output.flush()?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn generated_source_has_requested_section_count() {
        for section_count in 1..=4 {
            let archive = source_archive(section_count).unwrap();
            let package = open_source_package(&archive).unwrap();
            let snapshot = package.section_layout_snapshot().unwrap();
            assert_eq!(snapshot.inventory().sections().len(), section_count);
        }
    }

    #[test]
    fn parser_requires_bounded_lifecycle_arguments() {
        let args = vec![
            OsString::from("--sections"),
            OsString::from("3"),
            OsString::from("--concurrent-commits"),
            OsString::from("2"),
            OsString::from("--policy"),
            OsString::from("exact"),
            OsString::from("--ownership"),
            OsString::from("shared_snapshot_clones"),
            OsString::from("--source-revision"),
            OsString::from("0123456789abcdef0123456789abcdef01234567"),
            OsString::from("--output"),
            OsString::from("retention.json"),
        ];
        assert_eq!(
            parse_config(&args).unwrap(),
            Config {
                sections: 3,
                concurrent_commits: 2,
                policy: RetentionPolicy::Exact,
                ownership: OwnershipMode::SharedSnapshotClones,
                source_revision: "0123456789abcdef0123456789abcdef01234567".to_owned(),
                output: PathBuf::from("retention.json"),
            }
        );
        assert!(parse_config(&args[..4]).is_err());
        assert!(parse_config(&[OsString::from("--sections"), OsString::from("0")]).is_err());
    }
}
