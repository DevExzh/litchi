//! Paired allocator evidence for retained source-backed merge history.
//!
//! The measured interval keeps every committed patch alive while it performs
//! the same forward and inverse replay sequence for each patch.  The current
//! binary measures the W4 deferred source-backed owner.  The pre-W4 comparison
//! is kept as an exact external overlay on the pinned e63 checkout; it is not
//! a feature branch in this current workspace.  Both workflows retain the
//! same fixture, replay oracles and allocator counters.

#[path = "support/counting_allocator.rs"]
mod allocator;

use std::error::Error;
use std::io::Write;
use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};
use serde::Serialize;
use sha2::{Digest, Sha256};

use litchi_perf_baseline::allocation_metrics;

type AnyError = Box<dyn Error>;
type AnyResult<T> = Result<T, AnyError>;

const SCHEMA: &str = "litchi.xlsx.source-merge-history-alloc.v1";
const DEFAULT_ITERATIONS: usize = 64;
const DEFAULT_WARMUPS: usize = 2;
const WORKSHEET_PART: &str = "/xl/worksheets/sheet1.xml";
const EXPECTED_MERGE_ARCHIVE_SHA256: &str =
    "151fed9651e6f88a1e7e17183c8dac1f4885b6a922756214295b0d7c828a589e";
const EXPECTED_MERGE_TARGET_SHA256: &str =
    "692d1a1b71bd6af7bffc8d008d28ba75c5068dc913e0452363950fbe09d5b605";
const EXPECTED_UNMERGE_ARCHIVE_SHA256: &str =
    "6329afb234f9f1ea073e37baa4ca9ab6a0bb559fd40ff46a94320d416296a03f";
const EXPECTED_UNMERGE_TARGET_SHA256: &str =
    "467a33b4a4635f43d4d1a582ed8f31390a0eb6b71c94fc13059de9e5c436798e";

#[derive(Debug, Clone, Copy)]
struct Arguments {
    iterations: usize,
    warmups: usize,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    api_variant: &'static str,
    operation_pattern: &'static str,
    history_steps: usize,
    warmup_steps: usize,
    fixture: FixtureIdentity,
    exact_replay_checks: usize,
    exact_inverse_checks: usize,
    allocator: AllocationDelta,
    verification_allocator: AllocationDelta,
    common_replay_exact_replay_checks: usize,
    common_replay_exact_inverse_checks: usize,
    common_replay_allocator: AllocationDelta,
    common_replay_apply_allocator: AllocationDelta,
    common_replay_apply_forward_allocator: AllocationDelta,
    common_replay_apply_inverse_allocator: AllocationDelta,
    common_replay_publication_allocator: AllocationDelta,
    common_replay_forward_allocator: AllocationDelta,
    common_replay_inverse_allocator: AllocationDelta,
    common_replay_publication_forward_allocator: AllocationDelta,
    common_replay_publication_inverse_allocator: AllocationDelta,
    source_owner_admission_allocator: AllocationDelta,
}

#[derive(Debug, Serialize)]
struct FixtureIdentity {
    merge_archive_sha256: String,
    merge_target_sha256: String,
    unmerge_archive_sha256: String,
    unmerge_target_sha256: String,
    merge_bytes: usize,
    unmerge_bytes: usize,
}

#[derive(Debug, Serialize)]
struct AllocationDelta {
    status: &'static str,
    scope: &'static str,
    allocation_calls: u64,
    deallocation_calls: u64,
    reallocation_calls: u64,
    allocated_bytes: u64,
    deallocated_bytes: u64,
    live_bytes_before: u64,
    live_bytes_after: u64,
    peak_live_bytes_before: u64,
    peak_live_bytes_after: u64,
}

const API_VARIANT: &str = "w4-deferred-source-workbook";

fn main() -> AnyResult<()> {
    allocation_metrics::enable();
    let arguments = parse_arguments(std::env::args().skip(1))?;
    let (merge_fixture, unmerge_fixture) = fixtures()?;
    let fixture = fixture_identity(&merge_fixture, &unmerge_fixture)?;
    verify_fixture(&fixture)?;

    let before = allocation_metrics::snapshot();
    let patches = run_history(
        arguments.iterations,
        arguments.warmups,
        &merge_fixture,
        &unmerge_fixture,
    )?;
    let retained = allocation_metrics::snapshot();
    let (exact_replay_checks, exact_inverse_checks) =
        verify_history(&patches, &merge_fixture, &unmerge_fixture)?;
    let after = allocation_metrics::snapshot();
    let (
        common_replay_exact_replay_checks,
        common_replay_exact_inverse_checks,
        common_replay_allocator,
        common_replay_apply_allocator,
        common_replay_apply_forward_allocator,
        common_replay_apply_inverse_allocator,
        common_replay_publication_allocator,
        common_replay_forward_allocator,
        common_replay_inverse_allocator,
        common_replay_publication_forward_allocator,
        common_replay_publication_inverse_allocator,
    ) = run_common_replay(&patches, &merge_fixture, &unmerge_fixture)?;
    let source_owner_admission_allocator = run_source_owner_admission(
        arguments.iterations,
        arguments.warmups,
        &merge_fixture,
        &unmerge_fixture,
    )?;
    let allocator = allocation_delta(before, retained, "retained_patch_history")?;
    let verification_allocator = allocation_delta(
        retained,
        after,
        "independent_recipient_replay_with_required_facade_admission",
    )?;

    if exact_replay_checks != arguments.iterations || exact_inverse_checks != arguments.iterations {
        return Err("the retained history did not pass every exact replay check".into());
    }
    // Keep the patches alive until after the snapshot above.  The explicit
    // length check prevents an optimizer or an accidental refactor from
    // making the retained-history interval empty.
    if patches.len() != arguments.iterations {
        return Err("retained patch history length changed unexpectedly".into());
    }

    let report = Report {
        schema: SCHEMA,
        api_variant: API_VARIANT,
        operation_pattern: "merge,unmerge repeated from independently reopened equal source",
        history_steps: arguments.iterations,
        warmup_steps: arguments.warmups,
        fixture,
        exact_replay_checks,
        exact_inverse_checks,
        allocator,
        verification_allocator,
        common_replay_exact_replay_checks,
        common_replay_exact_inverse_checks,
        common_replay_allocator,
        common_replay_apply_allocator,
        common_replay_apply_forward_allocator,
        common_replay_apply_inverse_allocator,
        common_replay_publication_allocator,
        common_replay_forward_allocator,
        common_replay_inverse_allocator,
        common_replay_publication_forward_allocator,
        common_replay_publication_inverse_allocator,
        source_owner_admission_allocator,
    };
    let stdout = std::io::stdout();
    let mut output = stdout.lock();
    serde_json::to_writer_pretty(&mut output, &report)?;
    output.write_all(b"\n")?;
    Ok(())
}

#[derive(Debug)]
struct StepChecks {
    patch: Patch,
}

type Patch = litchi_xlsx::merge::source::Patch;

fn run_history(
    iterations: usize,
    warmups: usize,
    merge_fixture: &[u8],
    unmerge_fixture: &[u8],
) -> AnyResult<Vec<Patch>> {
    use litchi_xlsx::SourceBackedWorkbook;

    let merge_owner =
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(merge_fixture.to_vec())))?;
    let unmerge_owner =
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(unmerge_fixture.to_vec())))?;
    for step in 0..warmups {
        if step % 2 == 0 {
            run_merge_step(&merge_owner)?;
        } else {
            run_unmerge_step(&unmerge_owner)?;
        }
    }
    let mut patches = Vec::with_capacity(iterations);
    for step in 0..iterations {
        let checks = if step % 2 == 0 {
            run_merge_step(&merge_owner)?
        } else {
            run_unmerge_step(&unmerge_owner)?
        };
        patches.push(checks.patch);
    }
    Ok(patches)
}

fn run_merge_step(owner: &litchi_xlsx::SourceBackedWorkbook) -> AnyResult<StepChecks> {
    let mut edit = owner
        .edit_merges("Sheet1")?
        .ok_or("merge fixture worksheet disappeared")?;
    edit.merge("A1:B2")?;
    let commit = edit.commit()?;
    let patch = commit.patch().clone();
    Ok(StepChecks { patch })
}

fn run_unmerge_step(owner: &litchi_xlsx::SourceBackedWorkbook) -> AnyResult<StepChecks> {
    let mut edit = owner
        .edit_merges("Sheet1")?
        .ok_or("unmerge fixture worksheet disappeared")?;
    edit.unmerge("B2")?;
    let commit = edit.commit()?;
    let patch = commit.patch().clone();
    Ok(StepChecks { patch })
}

fn verify_history(
    patches: &[Patch],
    merge_fixture: &[u8],
    unmerge_fixture: &[u8],
) -> AnyResult<(usize, usize)> {
    use litchi_xlsx::SourceBackedWorkbook;

    let mut replay_checks = 0usize;
    let mut inverse_checks = 0usize;
    for (step, patch) in patches.iter().enumerate() {
        let fixture = if step % 2 == 0 {
            merge_fixture
        } else {
            unmerge_fixture
        };
        let replay_owner =
            SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(fixture.to_vec())))?;
        let replayed = replay_owner.apply_merges(patch)?.package().to_bytes()?;
        if replayed == fixture {
            return Err(format!("step {step} exact forward replay was a no-op").into());
        }
        replay_checks = replay_checks
            .checked_add(1)
            .ok_or("replay check counter overflow")?;
        let inverse_owner =
            SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(replayed)))?;
        let restored = inverse_owner
            .apply_merges(&patch.inverse())?
            .package()
            .to_bytes()?;
        if restored != fixture {
            return Err(format!("step {step} exact inverse replay changed bytes").into());
        }
        inverse_checks = inverse_checks
            .checked_add(1)
            .ok_or("inverse check counter overflow")?;
    }
    Ok((replay_checks, inverse_checks))
}

fn run_common_replay(
    patches: &[Patch],
    merge_fixture: &[u8],
    unmerge_fixture: &[u8],
) -> AnyResult<(
    usize,
    usize,
    AllocationDelta,
    AllocationDelta,
    AllocationDelta,
    AllocationDelta,
    AllocationDelta,
    AllocationDelta,
    AllocationDelta,
    AllocationDelta,
    AllocationDelta,
)> {
    use litchi_xlsx::SourceBackedWorkbook;

    let merge_recipient =
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(merge_fixture.to_vec())))?;
    let unmerge_recipient =
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(unmerge_fixture.to_vec())))?;
    let merge_forward_commit = merge_recipient.apply_merges(&patches[0])?;
    let merge_forward = merge_forward_commit.package().to_bytes()?;
    let merge_forward_pending = merge_forward_commit.package().clone();
    let merge_forward_recipient =
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(merge_forward.clone())))?;
    let unmerge_forward = if let Some(patch) = patches.get(1) {
        let commit = unmerge_recipient.apply_merges(patch)?;
        let output = commit.package().to_bytes()?;
        Some((
            output.clone(),
            commit.package().clone(),
            SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(output)))?,
        ))
    } else {
        None
    };

    let merge_inverse_recipient = merge_forward_recipient
        .apply_merges(&patches[0].inverse())?
        .package()
        .clone();
    let unmerge_inverse_recipient = if let Some((_, _, recipient)) = unmerge_forward.as_ref() {
        Some(
            recipient
                .apply_merges(&patches[1].inverse())?
                .package()
                .clone(),
        )
    } else {
        None
    };

    let apply_before = allocation_metrics::snapshot();
    for (step, patch) in patches.iter().enumerate() {
        let recipient = if step % 2 == 0 {
            merge_recipient.clone()
        } else {
            unmerge_recipient.clone()
        };
        let _ = recipient.apply_merges(patch)?;
    }
    let apply_forward_after = allocation_metrics::snapshot();
    for (step, patch) in patches.iter().enumerate() {
        let recipient = if step % 2 == 0 {
            merge_forward_recipient.clone()
        } else {
            unmerge_forward
                .as_ref()
                .ok_or("missing unmerge forward recipient")?
                .2
                .clone()
        };
        let _ = recipient.apply_merges(&patch.inverse())?;
    }
    let apply_after = allocation_metrics::snapshot();
    let apply_allocator = allocation_delta(
        apply_before,
        apply_after,
        "common_replay_apply_without_publication",
    )?;
    let apply_forward_allocator = allocation_delta(
        apply_before,
        apply_forward_after,
        "common_replay_apply_forward_without_publication",
    )?;
    let apply_inverse_allocator = allocation_delta(
        apply_forward_after,
        apply_after,
        "common_replay_apply_inverse_without_publication",
    )?;

    let publication_before = allocation_metrics::snapshot();
    let mut publication_replay_checks = 0usize;
    let mut publication_inverse_checks = 0usize;
    let publication_forward_before = allocation_metrics::snapshot();
    for (step, _) in patches.iter().enumerate() {
        let fixture = if step % 2 == 0 {
            merge_fixture
        } else {
            unmerge_fixture
        };
        let expected_forward = if step % 2 == 0 {
            &merge_forward
        } else {
            &unmerge_forward
                .as_ref()
                .ok_or("missing unmerge forward recipient")?
                .0
        };
        let forward = if step % 2 == 0 {
            merge_forward_pending.clone().to_bytes()?
        } else {
            unmerge_forward
                .as_ref()
                .ok_or("missing unmerge forward recipient")?
                .1
                .clone()
                .to_bytes()?
        };
        if forward.as_slice() != expected_forward.as_slice() || forward == fixture {
            return Err(format!(
                "step {step} common publication forward changed bytes unexpectedly"
            )
            .into());
        }
        publication_replay_checks = publication_replay_checks
            .checked_add(1)
            .ok_or("common publication replay check counter overflow")?;
    }
    let publication_forward_after = allocation_metrics::snapshot();
    let publication_inverse_before = allocation_metrics::snapshot();
    for (step, _) in patches.iter().enumerate() {
        let fixture = if step % 2 == 0 {
            merge_fixture
        } else {
            unmerge_fixture
        };

        let restored = if step % 2 == 0 {
            merge_inverse_recipient.clone().to_bytes()?
        } else {
            unmerge_inverse_recipient
                .as_ref()
                .ok_or("missing unmerge inverse recipient")?
                .clone()
                .to_bytes()?
        };
        if restored != fixture {
            return Err(format!(
                "step {step} common publication inverse changed bytes unexpectedly"
            )
            .into());
        }
        publication_inverse_checks = publication_inverse_checks
            .checked_add(1)
            .ok_or("common publication inverse check counter overflow")?;
    }
    let publication_inverse_after = allocation_metrics::snapshot();
    let publication_after = allocation_metrics::snapshot();
    if publication_replay_checks != patches.len() || publication_inverse_checks != patches.len() {
        return Err("common publication checks did not cover every patch".into());
    }
    let publication_allocator = allocation_delta(
        publication_before,
        publication_after,
        "common_replay_publication_without_apply",
    )?;
    let publication_forward_allocator = allocation_delta(
        publication_forward_before,
        publication_forward_after,
        "common_replay_publication_forward",
    )?;
    let publication_inverse_allocator = allocation_delta(
        publication_inverse_before,
        publication_inverse_after,
        "common_replay_publication_inverse",
    )?;

    let before = allocation_metrics::snapshot();
    let mut replay_checks = 0usize;
    let mut inverse_checks = 0usize;
    let forward_before = allocation_metrics::snapshot();
    for (step, patch) in patches.iter().enumerate() {
        let fixture = if step % 2 == 0 {
            merge_fixture
        } else {
            unmerge_fixture
        };
        let expected_forward = if step % 2 == 0 {
            &merge_forward
        } else {
            &unmerge_forward
                .as_ref()
                .ok_or("missing unmerge forward recipient")?
                .0
        };
        let recipient = if step % 2 == 0 {
            merge_recipient.clone()
        } else {
            unmerge_recipient.clone()
        };
        let forward = recipient.apply_merges(patch)?.package().to_bytes()?;
        if forward.as_slice() != expected_forward.as_slice() || forward == fixture {
            return Err(
                format!("step {step} common forward replay changed bytes unexpectedly").into(),
            );
        }
        replay_checks = replay_checks
            .checked_add(1)
            .ok_or("common replay check counter overflow")?;
    }
    let forward_after = allocation_metrics::snapshot();
    let inverse_before = allocation_metrics::snapshot();
    for (step, patch) in patches.iter().enumerate() {
        let fixture = if step % 2 == 0 {
            merge_fixture
        } else {
            unmerge_fixture
        };
        let inverse_recipient = if step % 2 == 0 {
            merge_forward_recipient.clone()
        } else {
            unmerge_forward
                .as_ref()
                .ok_or("missing unmerge forward recipient")?
                .2
                .clone()
        };
        let restored = inverse_recipient
            .apply_merges(&patch.inverse())?
            .package()
            .to_bytes()?;
        if restored != fixture {
            return Err(format!("step {step} common inverse replay changed bytes").into());
        }
        inverse_checks = inverse_checks
            .checked_add(1)
            .ok_or("common inverse check counter overflow")?;
    }
    let inverse_after = allocation_metrics::snapshot();
    let after = allocation_metrics::snapshot();
    let common_forward_allocator =
        allocation_delta(forward_before, forward_after, "common_replay_forward")?;
    let common_inverse_allocator =
        allocation_delta(inverse_before, inverse_after, "common_replay_inverse")?;
    Ok((
        replay_checks,
        inverse_checks,
        allocation_delta(before, after, "common_replay_without_owner_admission")?,
        apply_allocator,
        apply_forward_allocator,
        apply_inverse_allocator,
        publication_allocator,
        common_forward_allocator,
        common_inverse_allocator,
        publication_forward_allocator,
        publication_inverse_allocator,
    ))
}

fn run_source_owner_admission(
    iterations: usize,
    warmups: usize,
    merge_fixture: &[u8],
    unmerge_fixture: &[u8],
) -> AnyResult<AllocationDelta> {
    use litchi_xlsx::SourceBackedWorkbook;

    for step in 0..warmups {
        let fixture = if step % 2 == 0 {
            merge_fixture
        } else {
            unmerge_fixture
        };
        let owner =
            SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(fixture.to_vec())))?;
        let _edit = owner.edit_merges("Sheet1")?;
    }
    let before = allocation_metrics::snapshot();
    for step in 0..iterations {
        let fixture = if step % 2 == 0 {
            merge_fixture
        } else {
            unmerge_fixture
        };
        let owner =
            SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(fixture.to_vec())))?;
        let _edit = owner.edit_merges("Sheet1")?;
    }
    let after = allocation_metrics::snapshot();
    allocation_delta(before, after, "source_owner_admission_with_selected_edit")
}

fn fixtures() -> AnyResult<(Vec<u8>, Vec<u8>)> {
    Ok((fixture(false)?, fixture(true)?))
}

fn fixture(merged: bool) -> AnyResult<Vec<u8>> {
    let workbook = litchi_xlsx::Package::new()?;
    let mut edit = workbook.edit()?;
    {
        let mut sheet = edit
            .sheet("Sheet1")?
            .ok_or("minimal workbook worksheet disappeared")?;
        sheet.set("A1", "litchi-xlsx-merge-anchor-v1")?;
        sheet.set("C1", "litchi-xlsx-merge-unrelated-v1")?;
        if merged {
            sheet.merge("A1:B2")?;
        }
    }
    let committed = edit.commit()?.into_package();
    let mut output = Vec::new();
    committed.write_to(&mut output)?;
    let mut package = OpcPackage::from_bytes(&output, litchi_opc::SourceLimits::default())?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/merge-abba-aux.bin")?,
        "application/octet-stream".to_owned(),
        b"litchi-xlsx-source-merge-aux-v1".to_vec(),
    )))?;
    Ok(PackageWriter::to_bytes(&package)?)
}

fn fixture_identity(merge: &[u8], unmerge: &[u8]) -> AnyResult<FixtureIdentity> {
    Ok(FixtureIdentity {
        merge_archive_sha256: sha256_hex(merge),
        merge_target_sha256: target_sha256(merge)?,
        unmerge_archive_sha256: sha256_hex(unmerge),
        unmerge_target_sha256: target_sha256(unmerge)?,
        merge_bytes: merge.len(),
        unmerge_bytes: unmerge.len(),
    })
}

fn target_sha256(fixture: &[u8]) -> AnyResult<String> {
    let package = OpcPackage::from_bytes(fixture, litchi_opc::SourceLimits::default())?;
    let worksheet = PackURI::new(WORKSHEET_PART)?;
    Ok(sha256_hex(package.get_part(&worksheet)?.blob()))
}

fn verify_fixture(identity: &FixtureIdentity) -> AnyResult<()> {
    if identity.merge_archive_sha256 != EXPECTED_MERGE_ARCHIVE_SHA256
        || identity.unmerge_archive_sha256 != EXPECTED_UNMERGE_ARCHIVE_SHA256
    {
        return Err("history fixture archive identity changed".into());
    }
    // The target digest is checked by the exact replay oracles as well.  Keep
    // the pinned values here so a copied baseline adapter cannot silently use
    // a different generated corpus.
    if identity.merge_target_sha256 != EXPECTED_MERGE_TARGET_SHA256
        || identity.unmerge_target_sha256 != EXPECTED_UNMERGE_TARGET_SHA256
    {
        return Err("history fixture target identity changed".into());
    }
    Ok(())
}

fn allocation_delta(
    before: allocation_metrics::Snapshot,
    after: allocation_metrics::Snapshot,
    scope: &'static str,
) -> AnyResult<AllocationDelta> {
    if after.observer_invalid || after.overflowed {
        return Err("allocator counters became invalid during retained history".into());
    }
    let subtract = |after: u64, before: u64, label: &str| -> AnyResult<u64> {
        after
            .checked_sub(before)
            .ok_or_else(|| format!("allocator {label} counter regressed").into())
    };
    Ok(AllocationDelta {
        status: "measured",
        scope,
        allocation_calls: subtract(
            after.allocation_calls,
            before.allocation_calls,
            "allocation",
        )?,
        deallocation_calls: subtract(
            after.deallocation_calls,
            before.deallocation_calls,
            "deallocation",
        )?,
        reallocation_calls: subtract(
            after.reallocation_calls,
            before.reallocation_calls,
            "reallocation",
        )?,
        allocated_bytes: subtract(after.allocated_bytes, before.allocated_bytes, "allocated")?,
        deallocated_bytes: subtract(
            after.deallocated_bytes,
            before.deallocated_bytes,
            "deallocated",
        )?,
        live_bytes_before: before.live_bytes,
        live_bytes_after: after.live_bytes,
        peak_live_bytes_before: before.peak_live_bytes,
        peak_live_bytes_after: after.peak_live_bytes,
    })
}

fn parse_arguments<I>(arguments: I) -> AnyResult<Arguments>
where
    I: IntoIterator<Item = String>,
{
    let mut parsed = Arguments {
        iterations: DEFAULT_ITERATIONS,
        warmups: DEFAULT_WARMUPS,
    };
    let mut arguments = arguments.into_iter();
    while let Some(argument) = arguments.next() {
        let (name, value) = argument
            .split_once('=')
            .map_or((argument.as_str(), None), |(name, value)| {
                (name, Some(value))
            });
        match name {
            "--iterations" => {
                let value = value
                    .map(str::to_owned)
                    .or_else(|| arguments.next())
                    .ok_or("--iterations requires a value")?;
                parsed.iterations = value.parse()?;
            },
            "--warmups" => {
                let value = value
                    .map(str::to_owned)
                    .or_else(|| arguments.next())
                    .ok_or("--warmups requires a value")?;
                parsed.warmups = value.parse()?;
            },
            "--help" | "-h" => {
                println!("xlsx_merge_history_alloc [--iterations N] [--warmups N]");
                std::process::exit(0);
            },
            other => return Err(format!("unknown argument '{other}'").into()),
        }
    }
    if parsed.iterations == 0 {
        return Err("--iterations must be positive".into());
    }
    Ok(parsed)
}

fn sha256_hex(bytes: &[u8]) -> String {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(64);
    for byte in digest {
        output.push(HEX[usize::from(byte >> 4)] as char);
        output.push(HEX[usize::from(byte & 0x0f)] as char);
    }
    output
}
