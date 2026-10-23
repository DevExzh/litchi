//! Paired allocator evidence for source-backed tab-state history.
//!
//! The fixture deliberately contains a non-empty workbook `definedNames`
//! catalog.  That matters because the W4 deferred tab-state snapshot retains
//! the catalog for the read-only workbook projections and for replay.  The
//! same source, edit sequence, exact forward/inverse checks, and allocation
//! interval run against the pre-W4 and W4 APIs through the two small adapters
//! below.

#[path = "support/counting_allocator.rs"]
mod allocator;

use std::error::Error;
use std::io::Write;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource,
};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, SourceCacheLimits, SourceLimits};
use serde::Serialize;
use sha2::{Digest, Sha256};

use litchi_perf_baseline::allocation_metrics;

type AnyError = Box<dyn Error>;
type AnyResult<T> = Result<T, AnyError>;

const SCHEMA: &str = "litchi.xlsx.tab-state-history-alloc.v1";
const DEFAULT_ITERATIONS: usize = 64;
const DEFAULT_WARMUPS: usize = 2;
const MAIN_PART: &str = "/xl/workbook.xml";
const SHEET_ONE: &str = "/xl/worksheets/sheet1.xml";
const SHEET_TWO: &str = "/xl/worksheets/sheet2.xml";
const MANAGED_MEMORY: u64 = 16 * 1024 * 1024;
const MANAGED_CACHE_BYTES: usize = 256 * 1024;
const MANAGED_CACHE_ENTRIES: usize = 4;
const CUSTOM_INPUT_BYTES: u64 = 8 * 1024 * 1024;
const CUSTOM_CACHE_BYTES: usize = 128 * 1024;
const CUSTOM_CACHE_ENTRIES: usize = 2;

// Keeping the pins here makes later baseline/candidate runs fail closed if
// the fixture drifts.
const EXPECTED_VISIBLE_ARCHIVE_SHA256: Option<&str> =
    Some("8f64158142f26aa49ecacda3264ba18f83f6e971d040667df52c61a3b52f69a4");
const EXPECTED_HIDDEN_ARCHIVE_SHA256: Option<&str> =
    Some("0267a1418c2bf3d08e4a1f906faf38a7d88658b88e9cb23480cc2931db252dc1");

#[derive(Debug, Clone, Copy)]
struct Arguments {
    iterations: usize,
    warmups: usize,
    print_fixture_identity: bool,
    profile: PolicyProfile,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PolicyProfile {
    Default,
    Managed,
    Custom,
}

impl PolicyProfile {
    const fn name(self) -> &'static str {
        match self {
            Self::Default => "default",
            Self::Managed => "managed",
            Self::Custom => "custom",
        }
    }
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    api_variant: &'static str,
    policy_profile: &'static str,
    operation_pattern: &'static str,
    catalog_defined_names: usize,
    history_steps: usize,
    warmup_steps: usize,
    fixture: FixtureIdentity,
    exact_forward_checks: usize,
    exact_inverse_checks: usize,
    forward_visible_sha256: String,
    forward_hidden_sha256: String,
    inverse_visible_sha256: String,
    inverse_hidden_sha256: String,
    retained_history_allocator: AllocationDelta,
    verification_allocator: AllocationDelta,
    source_owner_admission_allocator: AllocationDelta,
}

#[derive(Debug, Serialize)]
struct FixtureIdentity {
    visible_archive_sha256: String,
    hidden_archive_sha256: String,
    visible_workbook_sha256: String,
    hidden_workbook_sha256: String,
    visible_bytes: usize,
    hidden_bytes: usize,
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

const API_VARIANT: &str = "w4-tab-state-deferred-source-workbook";

#[derive(Debug)]
struct StepChecks {
    patch: Patch,
}

type Patch = litchi_xlsx::tab_state::source::Patch;

#[derive(Debug)]
struct Verification {
    exact_forward_checks: usize,
    exact_inverse_checks: usize,
    forward_visible_sha256: String,
    forward_hidden_sha256: String,
    inverse_visible_sha256: String,
    inverse_hidden_sha256: String,
}

fn custom_source_limits() -> AnyResult<SourceLimits> {
    Ok(SourceLimits::builder()
        .max_input_bytes(CUSTOM_INPUT_BYTES)?
        .build()?)
}

fn cache_limits(profile: PolicyProfile) -> AnyResult<SourceCacheLimits> {
    let (bytes, entries) = match profile {
        PolicyProfile::Default => (MANAGED_CACHE_BYTES, MANAGED_CACHE_ENTRIES),
        PolicyProfile::Managed => (MANAGED_CACHE_BYTES, MANAGED_CACHE_ENTRIES),
        PolicyProfile::Custom => (CUSTOM_CACHE_BYTES, CUSTOM_CACHE_ENTRIES),
    };
    Ok(SourceCacheLimits::new(bytes, entries)?)
}

fn managed_context() -> AnyResult<ExecutionContext> {
    let budget = Budget::root(
        "xlsx-tab-state-history-alloc",
        Limits::new(
            MANAGED_MEMORY,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).ok_or("managed object limit is zero")?,
        NonZeroUsize::new(1).ok_or("managed depth limit is zero")?,
        NonZeroU64::new(MANAGED_MEMORY).ok_or("managed memory limit is zero")?,
        0,
    )?;
    Ok(ExecutionContext::new(
        budget,
        cancellation,
        execution_limits,
    ))
}

fn open_workbook(
    fixture: &[u8],
    profile: PolicyProfile,
) -> AnyResult<litchi_xlsx::SourceBackedWorkbook> {
    let source = Arc::new(OwnedSource::new(fixture.to_vec()));
    let mut options = litchi_xlsx::OpenOptions::new();
    match profile {
        PolicyProfile::Default => {},
        PolicyProfile::Managed => {
            options = options
                .cache_limits(cache_limits(profile)?)
                .execution_context(managed_context()?);
        },
        PolicyProfile::Custom => {
            options = options
                .source_limits(custom_source_limits()?)
                .cache_limits(cache_limits(profile)?);
        },
    }
    Ok(litchi_xlsx::SourceBackedWorkbook::from_read_at_with_options(source, options)?)
}

fn main() -> AnyResult<()> {
    allocation_metrics::enable();
    let arguments = parse_arguments(std::env::args().skip(1))?;
    let (visible_fixture, hidden_fixture) = fixtures()?;
    let fixture = fixture_identity(&visible_fixture, &hidden_fixture)?;
    verify_fixture(&fixture, arguments.print_fixture_identity)?;

    let before = allocation_metrics::snapshot();
    let patches = run_history(
        arguments.iterations,
        arguments.warmups,
        arguments.profile,
        &visible_fixture,
        &hidden_fixture,
    )?;
    let retained = allocation_metrics::snapshot();
    let verification = verify_history(
        &patches,
        arguments.profile,
        &visible_fixture,
        &hidden_fixture,
    )?;
    let after = allocation_metrics::snapshot();
    let source_owner_admission_allocator = run_source_owner_admission(
        arguments.iterations,
        arguments.warmups,
        arguments.profile,
        &visible_fixture,
        &hidden_fixture,
    )?;

    if verification.exact_forward_checks != arguments.iterations
        || verification.exact_inverse_checks != arguments.iterations
    {
        return Err("tab-state history did not pass every exact replay check".into());
    }
    if patches.len() != arguments.iterations {
        return Err("retained tab-state patch history length changed unexpectedly".into());
    }

    let report = Report {
        schema: SCHEMA,
        api_variant: API_VARIANT,
        policy_profile: arguments.profile.name(),
        operation_pattern: "hide/show repeated from independently reopened equal sources with two defined names",
        catalog_defined_names: 2,
        history_steps: arguments.iterations,
        warmup_steps: arguments.warmups,
        fixture,
        exact_forward_checks: verification.exact_forward_checks,
        exact_inverse_checks: verification.exact_inverse_checks,
        forward_visible_sha256: verification.forward_visible_sha256,
        forward_hidden_sha256: verification.forward_hidden_sha256,
        inverse_visible_sha256: verification.inverse_visible_sha256,
        inverse_hidden_sha256: verification.inverse_hidden_sha256,
        retained_history_allocator: allocation_delta(
            before,
            retained,
            "retained_tab_state_patch_history_with_defined_names",
        )?,
        verification_allocator: allocation_delta(
            retained,
            after,
            "independent_tab_state_replay_with_defined_name_catalog",
        )?,
        source_owner_admission_allocator,
    };
    let stdout = std::io::stdout();
    let mut output = stdout.lock();
    serde_json::to_writer_pretty(&mut output, &report)?;
    output.write_all(b"\n")?;
    Ok(())
}

fn run_history(
    iterations: usize,
    warmups: usize,
    profile: PolicyProfile,
    visible_fixture: &[u8],
    hidden_fixture: &[u8],
) -> AnyResult<Vec<Patch>> {
    let visible_owner = open_workbook(visible_fixture, profile)?;
    let hidden_owner = open_workbook(hidden_fixture, profile)?;
    for step in 0..warmups {
        if step % 2 == 0 {
            run_hide_step(&visible_owner)?;
        } else {
            run_show_step(&hidden_owner)?;
        }
    }
    let mut patches = Vec::with_capacity(iterations);
    for step in 0..iterations {
        let checks = if step % 2 == 0 {
            run_hide_step(&visible_owner)?
        } else {
            run_show_step(&hidden_owner)?
        };
        patches.push(checks.patch);
    }
    Ok(patches)
}

fn run_hide_step(owner: &litchi_xlsx::SourceBackedWorkbook) -> AnyResult<StepChecks> {
    let mut edit = owner.edit_tab_state()?;
    if !edit.hide("Sheet2")? {
        return Err("visible tab-state fixture did not change on hide".into());
    }
    let commit = edit.commit()?;
    Ok(StepChecks {
        patch: commit.patch().clone(),
    })
}

fn run_show_step(owner: &litchi_xlsx::SourceBackedWorkbook) -> AnyResult<StepChecks> {
    let mut edit = owner.edit_tab_state()?;
    if !edit.show("Sheet2")? {
        return Err("hidden tab-state fixture did not change on show".into());
    }
    let commit = edit.commit()?;
    Ok(StepChecks {
        patch: commit.patch().clone(),
    })
}

fn verify_history(
    patches: &[Patch],
    profile: PolicyProfile,
    visible_fixture: &[u8],
    hidden_fixture: &[u8],
) -> AnyResult<Verification> {
    let mut verification = Verification {
        exact_forward_checks: 0,
        exact_inverse_checks: 0,
        forward_visible_sha256: String::new(),
        forward_hidden_sha256: String::new(),
        inverse_visible_sha256: String::new(),
        inverse_hidden_sha256: String::new(),
    };
    for (step, patch) in patches.iter().enumerate() {
        let fixture = if step % 2 == 0 {
            visible_fixture
        } else {
            hidden_fixture
        };
        let replay_owner = open_workbook(fixture, profile)?;
        let forward = replay_owner.apply_tab_state(patch)?.package().to_bytes()?;
        if forward == fixture {
            return Err(format!("step {step} exact forward replay was a no-op").into());
        }
        if step % 2 == 0 {
            verification.forward_visible_sha256 = sha256_hex(&forward);
        } else {
            verification.forward_hidden_sha256 = sha256_hex(&forward);
        }
        verification.exact_forward_checks = verification
            .exact_forward_checks
            .checked_add(1)
            .ok_or("forward check counter overflow")?;

        let inverse_owner = open_workbook(&forward, profile)?;
        let restored = inverse_owner
            .apply_tab_state(&patch.inverse())?
            .package()
            .to_bytes()?;
        if restored != fixture {
            return Err(format!("step {step} exact inverse replay changed bytes").into());
        }
        if step % 2 == 0 {
            verification.inverse_visible_sha256 = sha256_hex(&restored);
        } else {
            verification.inverse_hidden_sha256 = sha256_hex(&restored);
        }
        verification.exact_inverse_checks = verification
            .exact_inverse_checks
            .checked_add(1)
            .ok_or("inverse check counter overflow")?;
    }
    Ok(verification)
}

fn run_source_owner_admission(
    iterations: usize,
    warmups: usize,
    profile: PolicyProfile,
    visible_fixture: &[u8],
    hidden_fixture: &[u8],
) -> AnyResult<AllocationDelta> {
    for step in 0..warmups {
        let fixture = if step % 2 == 0 {
            visible_fixture
        } else {
            hidden_fixture
        };
        let owner = open_workbook(fixture, profile)?;
        let _edit = owner.edit_tab_state()?;
    }
    let before = allocation_metrics::snapshot();
    for step in 0..iterations {
        let fixture = if step % 2 == 0 {
            visible_fixture
        } else {
            hidden_fixture
        };
        let owner = open_workbook(fixture, profile)?;
        let _edit = owner.edit_tab_state()?;
    }
    let after = allocation_metrics::snapshot();
    allocation_delta(
        before,
        after,
        "tab_state_source_owner_admission_with_defined_names",
    )
}

fn fixtures() -> AnyResult<(Vec<u8>, Vec<u8>)> {
    Ok((fixture(false)?, fixture(true)?))
}

fn fixture(hidden: bool) -> AnyResult<Vec<u8>> {
    let state = if hidden { " state=\"hidden\"" } else { "" };
    let workbook = format!(
        r#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0" firstSheet="0"/></bookViews><definedNames><definedName name="Rate">Sheet1!$A$1</definedName><definedName name="Scoped" localSheetId="1">Sheet2!$B$2</definedName></definedNames><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdOne"/><sheet name="Sheet2" sheetId="2"{state} r:id="rIdTwo"/></sheets><calcPr calcId="7"/></workbook>"#
    );
    let sheet_one = br#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews><sheetData/></worksheet>"#;
    let sheet_two = br#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/></worksheet>"#;
    let mut package = OpcPackage::new();
    for (name, content_type, bytes) in [
        (
            MAIN_PART,
            litchi_opc::constants::content_type::SML_SHEET_MAIN,
            workbook.into_bytes(),
        ),
        (
            SHEET_ONE,
            litchi_opc::constants::content_type::SML_WORKSHEET,
            sheet_one.to_vec(),
        ),
        (
            SHEET_TWO,
            litchi_opc::constants::content_type::SML_WORKSHEET,
            sheet_two.to_vec(),
        ),
    ] {
        package.try_add_part(Box::new(BlobPart::new(
            PackURI::new(name)?,
            content_type.to_owned(),
            bytes,
        )))?;
    }
    let main = PackURI::new(MAIN_PART)?;
    let workbook_part = package.get_part_mut(&main)?;
    for (target, id) in [
        ("worksheets/sheet1.xml", "rIdOne"),
        ("worksheets/sheet2.xml", "rIdTwo"),
    ] {
        workbook_part.rels_mut().try_add_relationship(
            litchi_opc::constants::relationship_type::WORKSHEET.to_owned(),
            target.to_owned(),
            id.to_owned(),
            litchi_opc::TargetMode::Internal,
        )?;
    }
    package.relate_to(
        "xl/workbook.xml",
        litchi_opc::constants::relationship_type::OFFICE_DOCUMENT,
    );
    Ok(PackageWriter::to_bytes(&package)?)
}

fn fixture_identity(visible: &[u8], hidden: &[u8]) -> AnyResult<FixtureIdentity> {
    Ok(FixtureIdentity {
        visible_archive_sha256: sha256_hex(visible),
        hidden_archive_sha256: sha256_hex(hidden),
        visible_workbook_sha256: workbook_sha256(visible)?,
        hidden_workbook_sha256: workbook_sha256(hidden)?,
        visible_bytes: visible.len(),
        hidden_bytes: hidden.len(),
    })
}

fn verify_fixture(identity: &FixtureIdentity, print: bool) -> AnyResult<()> {
    if print {
        eprintln!(
            "visible_archive_sha256={} hidden_archive_sha256={} visible_workbook_sha256={} hidden_workbook_sha256={} visible_bytes={} hidden_bytes={}",
            identity.visible_archive_sha256,
            identity.hidden_archive_sha256,
            identity.visible_workbook_sha256,
            identity.hidden_workbook_sha256,
            identity.visible_bytes,
            identity.hidden_bytes,
        );
    }
    if let Some(expected) = EXPECTED_VISIBLE_ARCHIVE_SHA256
        && identity.visible_archive_sha256 != expected
    {
        return Err("visible tab-state fixture archive identity changed".into());
    }
    if let Some(expected) = EXPECTED_HIDDEN_ARCHIVE_SHA256
        && identity.hidden_archive_sha256 != expected
    {
        return Err("hidden tab-state fixture archive identity changed".into());
    }
    Ok(())
}

fn workbook_sha256(bytes: &[u8]) -> AnyResult<String> {
    let package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default())?;
    Ok(sha256_hex(
        package.get_part(&PackURI::new(MAIN_PART)?)?.blob(),
    ))
}

fn allocation_delta(
    before: allocation_metrics::Snapshot,
    after: allocation_metrics::Snapshot,
    scope: &'static str,
) -> AnyResult<AllocationDelta> {
    if after.observer_invalid || after.overflowed {
        return Err("allocator counters became invalid during tab-state history".into());
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
        print_fixture_identity: false,
        profile: PolicyProfile::Default,
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
            "--profile" => {
                let value = value
                    .map(str::to_owned)
                    .or_else(|| arguments.next())
                    .ok_or("--profile requires a value")?;
                parsed.profile = match value.as_str() {
                    "default" => PolicyProfile::Default,
                    "managed" => PolicyProfile::Managed,
                    "custom" => PolicyProfile::Custom,
                    other => return Err(format!("unknown policy profile '{other}'").into()),
                };
            },
            "--print-fixture-identity" => parsed.print_fixture_identity = true,
            "--help" | "-h" => {
                println!(
                    "xlsx_tab_state_history_alloc [--iterations N] [--warmups N] [--profile default|managed|custom] [--print-fixture-identity]"
                );
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
