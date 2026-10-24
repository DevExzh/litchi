//! Real source-backed DrawingML adapters for the paired raw qualification.
//!
//! This endpoint exercises the shared `litchi-drawingml` codecs directly and
//! leaves package placement to the owning format crates. Every qualified row
//! carries an outside-clock fidelity receipt. A missing adapter or typed
//! refusal is a hard error; it is never converted into a timing claim.

#![forbid(unsafe_code)]

use std::env;
use std::fmt;
use std::fs;
use std::hint::black_box;
use std::io::{self, Write};
use std::path::{Path, PathBuf};
use std::time::{Duration, Instant};

use litchi_drawingml as dml;
use quick_xml::Reader;
use quick_xml::events::Event;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

const LANES: &[&str] = &[
    "drawingml-chart-raw",
    "drawingml-diagram-raw",
    "drawingml-text-media-raw",
    "drawingml-theme-ink",
    "drawingml-geometry-color-formatcode2-transform",
];

const RAW_OPERATIONS: &[&str] = &[
    "entry_read",
    "read_shared_when_supported",
    "vec_write",
    "write_to_bounded_sink",
    "preflight_refusal",
    "malformed_refusal",
];

const OWNER_OPERATIONS: &[&str] = &[
    "noop_commit",
    "changed_commit",
    "forward_replay",
    "inverse_replay",
    "stale_replay",
];

#[derive(Debug)]
struct Arguments {
    lane: String,
    endpoint: String,
    owner_route: Option<String>,
    raw_scope: Option<String>,
    role: String,
    mode: String,
    fixture: PathBuf,
    source_revision: String,
    iterations: usize,
    warmups: usize,
    blocks: usize,
    samples: usize,
    target_ms: u64,
    minimum_ms: u64,
    output: PathBuf,
}

#[derive(Debug, Clone, Deserialize)]
struct FixtureLane {
    root: String,
    mutation: String,
    malformed: String,
}

#[derive(Debug, Deserialize)]
struct FixtureFile {
    lanes: std::collections::BTreeMap<String, FixtureLane>,
}

#[derive(Debug, Clone, Copy)]
pub struct AllocationToken {
    pub allocation_calls: u64,
    pub deallocation_calls: u64,
    pub reallocation_calls: u64,
    pub failed_calls: u64,
    pub direct_bytes: u64,
    pub deallocated_bytes: u64,
    pub realloc_old_bytes: u64,
    pub realloc_new_bytes: u64,
    pub live_bytes: u64,
    pub peak_live_bytes: u64,
}

#[derive(Debug, Clone, Copy, Default)]
pub struct AllocationMetrics {
    pub allocation_calls: u64,
    pub deallocation_calls: u64,
    pub reallocation_calls: u64,
    pub failed_calls: u64,
    pub direct_bytes: u64,
    pub deallocated_bytes: u64,
    pub realloc_old_bytes: u64,
    pub realloc_new_bytes: u64,
    pub live_before: u64,
    pub live_after: u64,
    pub peak_live_delta: u64,
    pub conservation_ok: bool,
    pub elapsed_claimable: bool,
}

/// A safe observation hook implemented only by the allocator executable.
pub trait AllocationObserver {
    fn begin(&self) -> AllocationToken;
    fn finish(&self, before: AllocationToken) -> AllocationMetrics;
}

#[derive(Debug, Clone, Copy)]
enum PartKind {
    Chart,
    FormatCode,
    Diagram,
    TextBody,
    SvgBlip,
    Model3d,
    Blip,
    Theme,
    ThemeFamily,
    ThemeFamilyPart,
    InkActions,
    Geometry,
    Color,
    Transform,
}

/// The owner route is deliberately separate from a shared codec route.  A
/// raw lane can therefore carry a real read/write qualification without
/// pretending that a byte comparison is a package edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum OwnerRoute {
    DrawingmlChart,
    DrawingmlThemeFamily,
    DrawingmlInkActions,
    DrawingmlTransform,
}

#[derive(Debug, Clone, Serialize)]
struct CapabilityReceipt {
    route: &'static str,
    kind: &'static str,
    status: &'static str,
    evidence: &'static str,
    read_shared: bool,
}

fn supports_shared(kind: PartKind) -> bool {
    matches!(
        kind,
        PartKind::FormatCode | PartKind::ThemeFamily | PartKind::ThemeFamilyPart
    )
}

const CHART_OWNER_ROUTES: &[OwnerRoute] = &[OwnerRoute::DrawingmlChart];
const DIAGRAM_OWNER_ROUTES: &[OwnerRoute] = &[];
const TEXT_MEDIA_OWNER_ROUTES: &[OwnerRoute] = &[];
const THEME_OWNER_ROUTES: &[OwnerRoute] = &[
    OwnerRoute::DrawingmlThemeFamily,
    OwnerRoute::DrawingmlInkActions,
];
const GEOMETRY_OWNER_ROUTES: &[OwnerRoute] = &[OwnerRoute::DrawingmlTransform];

fn owner_routes(lane: &str) -> &'static [OwnerRoute] {
    match lane {
        "drawingml-chart-raw" => CHART_OWNER_ROUTES,
        "drawingml-diagram-raw" => DIAGRAM_OWNER_ROUTES,
        "drawingml-text-media-raw" => TEXT_MEDIA_OWNER_ROUTES,
        "drawingml-theme-ink" => THEME_OWNER_ROUTES,
        "drawingml-geometry-color-formatcode2-transform" => GEOMETRY_OWNER_ROUTES,
        _ => &[],
    }
}

fn owner_route_name(route: OwnerRoute) -> &'static str {
    match route {
        OwnerRoute::DrawingmlChart => "litchi-drawingml::chart::edit::Package::edit",
        OwnerRoute::DrawingmlThemeFamily => "litchi-drawingml::theme::family::Package::edit",
        OwnerRoute::DrawingmlInkActions => "litchi-drawingml::ink::actions::Prepared/Edit",
        OwnerRoute::DrawingmlTransform => "litchi-drawingml::transform::Package::edit",
    }
}

fn owner_route_from_name(value: &str) -> Option<OwnerRoute> {
    match value {
        "drawingml-chart" | "litchi-drawingml::chart::edit" => Some(OwnerRoute::DrawingmlChart),
        "drawingml-theme-family" | "litchi-drawingml::theme::family" => {
            Some(OwnerRoute::DrawingmlThemeFamily)
        },
        "drawingml-ink-actions" | "litchi-drawingml::ink::actions" => {
            Some(OwnerRoute::DrawingmlInkActions)
        },
        "drawingml-transform" | "litchi-drawingml::transform" => {
            Some(OwnerRoute::DrawingmlTransform)
        },
        _ => None,
    }
}

fn owner_route_evidence(route: OwnerRoute) -> &'static str {
    match route {
        OwnerRoute::DrawingmlChart => {
            "crates/litchi-drawingml/src/chart/edit/transaction.rs:Package::edit,Edit::commit,Package::apply"
        },
        OwnerRoute::DrawingmlThemeFamily => {
            "crates/litchi-drawingml/src/theme/family/transaction.rs:Package::edit,Edit::commit,Package::apply"
        },
        OwnerRoute::DrawingmlInkActions => {
            "crates/litchi-drawingml/src/ink/actions_edit.rs:Prepared + Edit::from_profile,Edit::commit,Patch::apply"
        },
        OwnerRoute::DrawingmlTransform => {
            "crates/litchi-drawingml/src/transform/transaction.rs:Package::edit,Edit::commit,Package::apply"
        },
    }
}

fn raw_route_evidence(kind: PartKind) -> &'static str {
    match kind {
        PartKind::Chart => "crates/litchi-drawingml/src/chart/reader/mod.rs;chart writer",
        PartKind::FormatCode => {
            "crates/litchi-drawingml/src/chart/extension/formatcode2.rs:read,write,write_to"
        },
        PartKind::Diagram => {
            "crates/litchi-drawingml/src/diagram/edit.rs:Outline::read,set_text,set_drawing_reference"
        },
        PartKind::TextBody => {
            "crates/litchi-drawingml/src/text/body/reader.rs:read_with_options;writer"
        },
        PartKind::SvgBlip => "crates/litchi-drawingml/src/svg_blip.rs:read,write,write_to",
        PartKind::Model3d => "crates/litchi-drawingml/src/model3d/codec/xml.rs:read,write,write_to",
        PartKind::Blip => "crates/litchi-drawingml/src/blip.rs:find_first_embed,write_embed",
        PartKind::Theme => "crates/litchi-drawingml/src/theme/codec.rs:read,encode_part",
        PartKind::ThemeFamily => {
            "crates/litchi-drawingml/src/theme/family/codec.rs:read,write,write_to"
        },
        PartKind::ThemeFamilyPart => {
            "crates/litchi-drawingml/src/theme/family/part.rs:read,add_family,replace_family,remove_family"
        },
        PartKind::InkActions => {
            "crates/litchi-drawingml/src/ink/actions.rs:read_profile,write_profile,write_profile_to"
        },
        PartKind::Geometry => "crates/litchi-drawingml/src/geometry/reader.rs;writer",
        PartKind::Color => "crates/litchi-drawingml/src/color/codec.rs:read,write",
        PartKind::Transform => "crates/litchi-drawingml/src/transform/codec.rs:read,write,write_to",
    }
}

fn capability_receipts(case: &LaneCase, endpoint: &str) -> Vec<CapabilityReceipt> {
    let mut receipts = Vec::new();
    if endpoint == "raw" {
        for part in &case.parts {
            receipts.push(CapabilityReceipt {
                route: part.id,
                kind: "raw",
                status: "qualified",
                evidence: raw_route_evidence(part.kind),
                read_shared: supports_shared(part.kind),
            });
        }
    } else {
        for fixture in &case.owners {
            for operation in OWNER_OPERATIONS {
                receipts.push(CapabilityReceipt {
                    route: operation,
                    kind: owner_route_name(fixture.route),
                    status: "qualified",
                    evidence: owner_route_evidence(fixture.route),
                    read_shared: false,
                });
            }
        }
    }
    receipts
}

#[derive(Debug, Clone)]
struct Part {
    id: &'static str,
    kind: PartKind,
    source: Vec<u8>,
    changed: Vec<u8>,
    malformed: Vec<u8>,
    over: Vec<u8>,
    limit_resource: &'static str,
}

#[derive(Debug, Clone)]
struct LaneCase {
    parts: Vec<Part>,
    stale: Vec<u8>,
    owners: Vec<OwnerFixture>,
    fixture_identity_sha256: String,
}

#[derive(Debug, Clone)]
struct OwnerFixture {
    route: OwnerRoute,
    source: Vec<u8>,
    stale: Vec<u8>,
}

#[derive(Debug)]
enum OwnerPatch {
    DrawingmlChart(dml::chart::edit::Patch),
    DrawingmlThemeFamily(dml::theme::family::Patch),
    DrawingmlInkActions(dml::ink::actions::Patch),
    DrawingmlTransform(dml::transform::Patch),
}

#[derive(Debug, Clone, Serialize)]
struct OwnerCommitReceipt {
    route: &'static str,
    fixture_identity_sha256: String,
    source_sha256: String,
    noop_sha256: String,
    stale_sha256: String,
    changed_sha256: String,
    forward_sha256: String,
    inverse_sha256: String,
    stale_kind: &'static str,
    stale_refusal: RefusalReceipt,
}

#[derive(Debug, Clone, Serialize)]
struct RawFidelityReceipt {
    part_id: &'static str,
    fixture_identity_sha256: String,
    source_sha256: String,
    source_output_sha256: String,
    changed_sha256: String,
    changed_output_sha256: String,
    malformed_sha256: String,
    changed_output_differs: bool,
}

#[derive(Debug, Clone, Serialize)]
struct ThemeFamilyMutationReceipt {
    operation: &'static str,
    source_sha256: String,
    output_sha256: String,
    output_bytes: usize,
    output_differs: bool,
}

#[derive(Debug, Clone, Serialize)]
struct RefusalReceipt {
    kind: &'static str,
    source_kind: Option<&'static str>,
    limit_kind: Option<&'static str>,
    resource: Option<&'static str>,
    limit: Option<usize>,
    observed: usize,
    message: String,
}

#[derive(Debug)]
struct ProbeFailure {
    message: String,
    refusal: Option<RefusalReceipt>,
}

#[derive(Debug, Clone, Serialize)]
struct SinkReceipt {
    writes: usize,
    bytes: usize,
    sha256: String,
}

#[derive(Debug, Clone, Serialize)]
struct SampleReceipt {
    elapsed_ns: u128,
    iterations: usize,
}

#[derive(Debug, Clone, Serialize)]
struct Row {
    operation: String,
    status: &'static str,
    elapsed_ns: u128,
    fidelity_checked: bool,
    typed_refusal_checked: bool,
    sink_writes: usize,
    sink: Option<SinkReceipt>,
    samples: Vec<SampleReceipt>,
    details: serde_json::Value,
}

#[derive(Debug, Clone, Serialize)]
struct SizeClassReceipt {
    part_id: &'static str,
    class: &'static str,
    bytes: usize,
    applicable: bool,
    accepted: bool,
    output_sha256: Option<String>,
    refusal: Option<RefusalReceipt>,
}

#[derive(Debug, Clone, Serialize)]
struct ResourceMetrics {
    process_rss_bytes: u64,
    peak_rss_bytes: u64,
    whole_process_peak_rss_bytes: u64,
}

#[derive(Debug, Clone, Serialize)]
struct AllocatorReport {
    elapsed_claimable: bool,
    allocation_calls: u64,
    deallocation_calls: u64,
    reallocation_calls: u64,
    failed_calls: u64,
    direct_bytes: u64,
    deallocated_bytes: u64,
    realloc_old_bytes: u64,
    realloc_new_bytes: u64,
    live_before: u64,
    live_after: u64,
    live_balance: i128,
    peak_live_delta: u64,
    process_rss_bytes: u64,
    peak_rss_bytes: u64,
    whole_process_peak_rss_bytes: u64,
    conservation_ok: bool,
}

#[derive(Debug, Clone, Serialize)]
struct Protocol {
    warmups: usize,
    samples_per_leg: usize,
    blocks: usize,
    target_ms: u64,
    minimum_ms: u64,
    fresh_process: bool,
    abba_order: [&'static str; 4],
}

#[derive(Debug, Clone, Serialize)]
struct Report {
    schema: &'static str,
    lane: String,
    endpoint: String,
    owner_route: Option<String>,
    raw_scope: Option<String>,
    role: String,
    mode: String,
    source_revision: String,
    status: &'static str,
    qualification_claim: bool,
    fixture_bytes: usize,
    fixture_sha256: String,
    fixture_root_sha256: String,
    fixture_mutation: String,
    fixture_lane_identity_sha256: String,
    capabilities: Vec<CapabilityReceipt>,
    raw_fidelity: Vec<RawFidelityReceipt>,
    theme_family_mutations: Vec<ThemeFamilyMutationReceipt>,
    owner_proofs: Vec<OwnerCommitReceipt>,
    size_classes: Vec<SizeClassReceipt>,
    protocol: Protocol,
    allocator_metrics: Option<AllocatorReport>,
    resource_metrics: ResourceMetrics,
    rows: Vec<Row>,
    iterations: usize,
}

fn usage() -> &'static str {
    "usage: drawingml-raw-{native,allocator} --lane LANE --role ROLE \
     --endpoint raw|owner --owner-route ROUTE --raw-scope SCOPE --mode MODE --fixture PATH --source-revision SHA \
     --iterations N --warmups N \
     --blocks N --samples N --target-ms N --minimum-ms N --json PATH"
}

fn invalid_argument(message: impl Into<String>) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidInput, message.into())
}

fn argument_value(arguments: &mut impl Iterator<Item = String>, name: &str) -> io::Result<String> {
    arguments
        .next()
        .ok_or_else(|| invalid_argument(format!("missing value for {name}")))
}

fn parse(arguments: impl Iterator<Item = String>, expected_mode: &str) -> io::Result<Arguments> {
    let mut lane = None;
    let mut endpoint = None;
    let mut owner_route = None;
    let mut raw_scope = None;
    let mut role = None;
    let mut mode = None;
    let mut fixture = None;
    let mut source_revision = None;
    let mut iterations = None;
    let mut warmups = None;
    let mut blocks = None;
    let mut samples = None;
    let mut target_ms = None;
    let mut minimum_ms = None;
    let mut output = None;
    let mut arguments = arguments;
    while let Some(argument) = arguments.next() {
        let slot = match argument.as_str() {
            "--lane" => &mut lane,
            "--endpoint" => &mut endpoint,
            "--owner-route" => &mut owner_route,
            "--raw-scope" => &mut raw_scope,
            "--role" => &mut role,
            "--mode" => &mut mode,
            "--fixture" => &mut fixture,
            "--source-revision" => &mut source_revision,
            "--iterations" => &mut iterations,
            "--warmups" => &mut warmups,
            "--blocks" => &mut blocks,
            "--samples" => &mut samples,
            "--target-ms" => &mut target_ms,
            "--minimum-ms" => &mut minimum_ms,
            "--json" => &mut output,
            _ => return Err(invalid_argument(format!("unknown argument {argument:?}"))),
        };
        if slot.is_some() {
            return Err(invalid_argument(format!("duplicate argument {argument}")));
        }
        *slot = Some(argument_value(&mut arguments, &argument)?);
    }

    let lane = lane.ok_or_else(|| invalid_argument("--lane is required"))?;
    if !LANES.contains(&lane.as_str()) {
        return Err(invalid_argument(format!(
            "unsupported DrawingML lane {lane:?}"
        )));
    }
    let endpoint = endpoint.ok_or_else(|| invalid_argument("--endpoint is required"))?;
    if !matches!(endpoint.as_str(), "raw" | "owner") {
        return Err(invalid_argument(format!(
            "unsupported endpoint {endpoint:?}"
        )));
    }
    let owner_route = owner_route;
    if endpoint == "owner" && owner_route.is_none() {
        return Err(invalid_argument("owner endpoint requires --owner-route"));
    }
    if endpoint == "raw" && owner_route.is_some() {
        return Err(invalid_argument("raw endpoint cannot carry --owner-route"));
    }
    if endpoint == "owner" && raw_scope.is_some() {
        return Err(invalid_argument("owner endpoint cannot carry --raw-scope"));
    }
    if let Some(scope) = raw_scope.as_deref() {
        if !matches!(scope, "bulk" | "family") {
            return Err(invalid_argument(format!("unsupported raw scope {scope:?}")));
        }
    }
    let role = role.ok_or_else(|| invalid_argument("--role is required"))?;
    if !matches!(role.as_str(), "baseline" | "candidate") {
        return Err(invalid_argument(format!("unsupported role {role:?}")));
    }
    let mode = mode.ok_or_else(|| invalid_argument("--mode is required"))?;
    if mode != expected_mode {
        return Err(invalid_argument(format!(
            "{expected_mode} binary requires --mode {expected_mode}"
        )));
    }
    let fixture = PathBuf::from(fixture.ok_or_else(|| invalid_argument("--fixture is required"))?);
    if fixture.is_symlink() || !fixture.is_file() {
        return Err(invalid_argument(format!(
            "fixture is not a regular file: {fixture:?}"
        )));
    }
    let source_revision =
        source_revision.ok_or_else(|| invalid_argument("--source-revision is required"))?;
    if source_revision.len() != 40
        || !source_revision
            .bytes()
            .all(|byte| byte.is_ascii_hexdigit() && !byte.is_ascii_uppercase())
    {
        return Err(invalid_argument(
            "--source-revision must be a lowercase 40-character SHA",
        ));
    }
    let parse_usize = |value: Option<String>, name: &str| -> io::Result<usize> {
        value
            .ok_or_else(|| invalid_argument(format!("{name} is required")))?
            .parse::<usize>()
            .map_err(|_| invalid_argument(format!("{name} must be an integer")))
    };
    let parse_u64 = |value: Option<String>, name: &str| -> io::Result<u64> {
        value
            .ok_or_else(|| invalid_argument(format!("{name} is required")))?
            .parse::<u64>()
            .map_err(|_| invalid_argument(format!("{name} must be an integer")))
    };
    let iterations = parse_usize(iterations, "--iterations")?;
    let warmups = parse_usize(warmups, "--warmups")?;
    let blocks = parse_usize(blocks, "--blocks")?;
    let samples = parse_usize(samples, "--samples")?;
    let target_ms = parse_u64(target_ms, "--target-ms")?;
    let minimum_ms = parse_u64(minimum_ms, "--minimum-ms")?;
    if iterations == 0 || warmups != 3 || blocks != 5 || samples != 20 {
        return Err(invalid_argument(
            "iterations must be positive, warmups=3, blocks=5, samples=20",
        ));
    }
    if target_ms < minimum_ms || minimum_ms < 150 {
        return Err(invalid_argument(
            "target-ms must be at least minimum-ms and minimum-ms must be >= 150",
        ));
    }
    let output = PathBuf::from(output.ok_or_else(|| invalid_argument("--json is required"))?);
    if output.as_os_str().is_empty() || output.is_symlink() || output.exists() {
        return Err(invalid_argument(format!(
            "JSON output must be a fresh regular-file path: {output:?}"
        )));
    }
    Ok(Arguments {
        lane,
        endpoint,
        owner_route,
        raw_scope,
        role,
        mode,
        fixture,
        source_revision,
        iterations,
        warmups,
        blocks,
        samples,
        target_ms,
        minimum_ms,
        output,
    })
}

fn sha256(bytes: &[u8]) -> String {
    let mut digest = Sha256::new();
    digest.update(bytes);
    format!("{:x}", digest.finalize())
}

fn load_fixture(path: &Path, lane: &str) -> io::Result<(Vec<u8>, FixtureLane)> {
    let bytes = fs::read(path)?;
    let value: FixtureFile = serde_json::from_slice(&bytes).map_err(|error| {
        invalid_argument(format!("fixture descriptor is invalid JSON: {error}"))
    })?;
    let lane_value = value
        .lanes
        .get(lane)
        .cloned()
        .ok_or_else(|| invalid_argument(format!("fixture descriptor has no lane {lane}")))?;
    Ok((bytes, lane_value))
}

fn fail(message: impl Into<String>) -> ProbeFailure {
    ProbeFailure {
        message: message.into(),
        refusal: None,
    }
}

fn invalid_input(message: impl Into<String>, observed: usize) -> ProbeFailure {
    let message = message.into();
    ProbeFailure {
        message: message.clone(),
        refusal: Some(RefusalReceipt {
            kind: "invalid_input",
            source_kind: None,
            limit_kind: None,
            resource: None,
            limit: None,
            observed,
            message,
        }),
    }
}

#[cfg(not(feature = "drawingml-owner-current"))]
fn classify(error: dml::Error, observed: usize) -> ProbeFailure {
    let message = error.to_string();
    let refusal = match error {
        dml::Error::Io(error) => Some(RefusalReceipt {
            kind: "io_error",
            source_kind: None,
            limit_kind: None,
            resource: None,
            limit: None,
            observed,
            message: error.to_string(),
        }),
        dml::Error::Limit { resource, limit } => Some(RefusalReceipt {
            kind: "limit",
            source_kind: None,
            limit_kind: None,
            resource: Some(resource),
            limit: Some(limit),
            observed,
            message: message.clone(),
        }),
        _ => Some(RefusalReceipt {
            kind: "invalid_input",
            source_kind: None,
            limit_kind: None,
            resource: None,
            limit: None,
            observed,
            message: message.clone(),
        }),
    };
    ProbeFailure { message, refusal }
}

#[cfg(feature = "drawingml-owner-current")]
fn classify(error: dml::Error, observed: usize) -> ProbeFailure {
    let message = error.to_string();
    let refusal = match error {
        dml::Error::Io(error) => Some(RefusalReceipt {
            kind: "io_error",
            source_kind: None,
            limit_kind: None,
            resource: None,
            limit: None,
            observed,
            message: error.to_string(),
        }),
        dml::Error::Limit {
            kind,
            observed: error_observed,
            maximum,
        } => Some(RefusalReceipt {
            kind: "limit",
            source_kind: None,
            limit_kind: Some(kind.resource()),
            resource: Some(kind.resource()),
            limit: Some(maximum),
            observed: error_observed.max(observed),
            message: message.clone(),
        }),
        dml::Error::PatchConflict => Some(RefusalReceipt {
            kind: "patch_conflict",
            source_kind: Some("patch_conflict"),
            limit_kind: None,
            resource: None,
            limit: None,
            observed,
            message: message.clone(),
        }),
        _ => Some(RefusalReceipt {
            kind: "invalid_input",
            source_kind: None,
            limit_kind: None,
            resource: None,
            limit: None,
            observed,
            message: message.clone(),
        }),
    };
    ProbeFailure { message, refusal }
}

fn io_failure(error: io::Error, observed: usize) -> ProbeFailure {
    ProbeFailure {
        message: error.to_string(),
        refusal: Some(RefusalReceipt {
            kind: "io_error",
            source_kind: None,
            limit_kind: None,
            resource: None,
            limit: None,
            observed,
            message: error.to_string(),
        }),
    }
}

fn read_chart(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let chart = dml::chart::reader::read(bytes).map_err(|error| classify(error, bytes.len()))?;
    let mut output = Vec::new();
    dml::chart::writer::write(&mut output, &chart)
        .map_err(|error| io_failure(error, bytes.len()))?;
    Ok(output)
}

fn read_format_code(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let value = dml::chart::extension::formatcode2::read(bytes)
        .map_err(|error| classify(error, bytes.len()))?;
    dml::chart::extension::formatcode2::write(&value).map_err(|error| classify(error, bytes.len()))
}

fn read_diagram(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let outline =
        dml::diagram::edit::Outline::read(bytes).map_err(|error| classify(error, bytes.len()))?;
    let entry = outline
        .entries
        .first()
        .ok_or_else(|| invalid_input("diagram data model has no content entry", bytes.len()))?;
    let output = dml::diagram::edit::set_text(bytes, &entry.id, &entry.text)
        .map_err(|error| classify(error, bytes.len()))?;
    dml::diagram::edit::set_drawing_reference(&output, outline.drawing_relationship.as_deref())
        .map_err(|error| classify(error, output.len()))
}

fn body_roundtrip(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    use dml::text::body::reader::Limits;
    let body = dml::text::body::reader::read_with_options(
        bytes,
        quick_xml::name::NamespaceResolver::default(),
        Limits::DEFAULT,
    )
    .map_err(|error| classify(error, bytes.len()))?;
    let mut contents = String::new();
    dml::text::body::writer::write(&mut contents, &body)
        .map_err(|error| classify(error, bytes.len()))?;
    let mut output = String::with_capacity(contents.len() + 32);
    output.push_str("<a:txBody xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">");
    output.push_str(&contents);
    output.push_str("</a:txBody>");
    Ok(output.into_bytes())
}

fn read_svg(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let value = dml::svg_blip::read(bytes).map_err(|error| classify(error, bytes.len()))?;
    dml::svg_blip::write(&value).map_err(|error| classify(error, bytes.len()))
}

fn read_model3d(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let value = dml::model3d::codec::read(bytes).map_err(|error| classify(error, bytes.len()))?;
    dml::model3d::codec::write(&value).map_err(|error| classify(error, bytes.len()))
}

fn read_blip(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let relationship =
        dml::blip::find_first_embed(bytes).map_err(|error| classify(error, bytes.len()))?;
    let Some(relationship) = relationship else {
        return Err(invalid_input(
            "image blip has no embedded relationship",
            bytes.len(),
        ));
    };
    let mut output = String::new();
    dml::blip::write_embed(&mut output, &relationship, true)
        .map_err(|error| invalid_input(error.to_string(), bytes.len()))?;
    Ok(output.into_bytes())
}

fn read_theme(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let theme = dml::theme::codec::read(bytes).map_err(|error| classify(error, bytes.len()))?;
    dml::theme::codec::encode_part(&theme.name, &theme.colors, &theme.fonts)
        .map_err(|error| classify(error, bytes.len()))
}

fn read_family(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let family = dml::theme::family::read(bytes).map_err(|error| classify(error, bytes.len()))?;
    dml::theme::family::write(&family).map_err(|error| classify(error, bytes.len()))
}

fn read_theme_family_part(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let snapshot =
        dml::theme::family::part::read(bytes).map_err(|error| classify(error, bytes.len()))?;
    let family = snapshot.family().ok_or_else(|| {
        invalid_input("theme family part has no direct family owner", bytes.len())
    })?;
    theme_family_replace(&snapshot, family).map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(feature = "drawingml-owner-current"))]
fn theme_family_replace(
    snapshot: &dml::theme::family::part::Snapshot,
    family: &dml::theme::family::Family,
) -> dml::Result<Vec<u8>> {
    snapshot.replace_family_with_limit(family, dml::theme::family::part::MAX_XML_BYTES)
}

#[cfg(feature = "drawingml-owner-current")]
fn theme_family_replace(
    snapshot: &dml::theme::family::part::Snapshot,
    family: &dml::theme::family::Family,
) -> dml::Result<Vec<u8>> {
    snapshot.replace_family(family, dml::theme::family::part::Limits::DEFAULT)
}

#[cfg(not(feature = "drawingml-owner-current"))]
fn theme_family_add(source: &[u8], family: &dml::theme::family::Family) -> dml::Result<Vec<u8>> {
    dml::theme::family::part::add_family_with_limit(
        source,
        family,
        dml::theme::family::part::MAX_XML_BYTES,
    )
}

#[cfg(feature = "drawingml-owner-current")]
fn theme_family_add(source: &[u8], family: &dml::theme::family::Family) -> dml::Result<Vec<u8>> {
    dml::theme::family::part::add_family(source, family, dml::theme::family::part::Limits::DEFAULT)
}

#[cfg(not(feature = "drawingml-owner-current"))]
fn theme_family_remove(source: &[u8]) -> dml::Result<Vec<u8>> {
    dml::theme::family::part::remove_family_with_limit(
        source,
        dml::theme::family::part::MAX_XML_BYTES,
    )
}

#[cfg(feature = "drawingml-owner-current")]
fn theme_family_remove(source: &[u8]) -> dml::Result<Vec<u8>> {
    dml::theme::family::part::remove_family(source, dml::theme::family::part::Limits::DEFAULT)
}

fn complete_theme_family_part(family: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let family = std::str::from_utf8(family).map_err(|error| {
        invalid_input(
            format!("theme family fragment is not UTF-8: {error}"),
            family.len(),
        )
    })?;
    Ok(format!(
        "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:thm15=\"{}\" name=\"Office\"><a:themeElements/><a:extLst><a:ext uri=\"{}\">{family}</a:ext></a:extLst></a:theme>",
        dml::theme::family::NAMESPACE,
        dml::theme::family::NAMESPACE,
    )
    .into_bytes())
}

fn read_ink(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let profile =
        dml::ink::actions::read_profile(bytes).map_err(|error| classify(error, bytes.len()))?;
    dml::ink::actions::write_profile(&profile).map_err(|error| classify(error, bytes.len()))
}

fn read_geometry(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let geometry = parse_geometry(bytes)?;
    let mut output = String::new();
    dml::geometry::writer::write_custom_geometry(&mut output, &geometry);
    Ok(output.into_bytes())
}

fn read_color(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let value = dml::color::read(bytes).map_err(|error| classify(error, bytes.len()))?;
    dml::color::write(&value).map_err(|error| classify(error, bytes.len()))
}

fn read_transform(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let value = dml::transform::codec::read(bytes).map_err(|error| classify(error, bytes.len()))?;
    dml::transform::codec::write(&value).map_err(|error| classify(error, bytes.len()))
}

fn roundtrip(kind: PartKind, bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    match kind {
        PartKind::Chart => read_chart(bytes),
        PartKind::FormatCode => read_format_code(bytes),
        PartKind::Diagram => read_diagram(bytes),
        PartKind::TextBody => body_roundtrip(bytes),
        PartKind::SvgBlip => read_svg(bytes),
        PartKind::Model3d => read_model3d(bytes),
        PartKind::Blip => read_blip(bytes),
        PartKind::Theme => read_theme(bytes),
        PartKind::ThemeFamily => read_family(bytes),
        PartKind::ThemeFamilyPart => read_theme_family_part(bytes),
        PartKind::InkActions => read_ink(bytes),
        PartKind::Geometry => read_geometry(bytes),
        PartKind::Color => read_color(bytes),
        PartKind::Transform => read_transform(bytes),
    }
}

fn write_kind_to<W: Write>(
    kind: PartKind,
    bytes: &[u8],
    writer: &mut W,
) -> Result<(), ProbeFailure> {
    match kind {
        PartKind::Chart => {
            let chart =
                dml::chart::reader::read(bytes).map_err(|error| classify(error, bytes.len()))?;
            dml::chart::writer::write(writer, &chart)
                .map_err(|error| io_failure(error, bytes.len()))
        },
        PartKind::FormatCode => {
            let value = dml::chart::extension::formatcode2::read(bytes)
                .map_err(|error| classify(error, bytes.len()))?;
            dml::chart::extension::formatcode2::write_to(writer, &value)
                .map_err(|error| classify(error, bytes.len()))
        },
        PartKind::Diagram => {
            let output = read_diagram(bytes)?;
            writer
                .write_all(&output)
                .map_err(|error| io_failure(error, output.len()))
        },
        PartKind::TextBody => write_text_body_to(bytes, writer),
        PartKind::SvgBlip => {
            let value = dml::svg_blip::read(bytes).map_err(|error| classify(error, bytes.len()))?;
            dml::svg_blip::write_to(writer, &value).map_err(|error| classify(error, bytes.len()))
        },
        PartKind::Model3d => {
            let value =
                dml::model3d::codec::read(bytes).map_err(|error| classify(error, bytes.len()))?;
            dml::model3d::codec::write_to(writer, &value)
                .map_err(|error| classify(error, bytes.len()))
        },
        PartKind::Blip => write_blip_to(bytes, writer),
        PartKind::Theme => {
            let value =
                dml::theme::codec::read(bytes).map_err(|error| classify(error, bytes.len()))?;
            let output = dml::theme::codec::encode_part(&value.name, &value.colors, &value.fonts)
                .map_err(|error| classify(error, bytes.len()))?;
            writer
                .write_all(&output)
                .map_err(|error| io_failure(error, bytes.len()))
        },
        PartKind::ThemeFamily => {
            let value =
                dml::theme::family::read(bytes).map_err(|error| classify(error, bytes.len()))?;
            dml::theme::family::write_to(writer, &value)
                .map_err(|error| classify(error, bytes.len()))
        },
        PartKind::ThemeFamilyPart => {
            let snapshot = dml::theme::family::part::read(bytes)
                .map_err(|error| classify(error, bytes.len()))?;
            let family = snapshot.family().ok_or_else(|| {
                invalid_input("theme family part has no direct family owner", bytes.len())
            })?;
            let output = theme_family_replace(&snapshot, family)
                .map_err(|error| classify(error, bytes.len()))?;
            writer
                .write_all(&output)
                .map_err(|error| io_failure(error, output.len()))
        },
        PartKind::InkActions => {
            let value = dml::ink::actions::read_profile(bytes)
                .map_err(|error| classify(error, bytes.len()))?;
            dml::ink::actions::write_profile_to(writer, &value)
                .map_err(|error| classify(error, bytes.len()))
        },
        PartKind::Geometry => write_geometry_to(bytes, writer),
        PartKind::Color => {
            let value = dml::color::read(bytes).map_err(|error| classify(error, bytes.len()))?;
            let output = dml::color::write(&value).map_err(|error| classify(error, bytes.len()))?;
            writer
                .write_all(&output)
                .map_err(|error| io_failure(error, bytes.len()))
        },
        PartKind::Transform => {
            let value =
                dml::transform::codec::read(bytes).map_err(|error| classify(error, bytes.len()))?;
            dml::transform::codec::write_to(writer, &value)
                .map_err(|error| classify(error, bytes.len()))
        },
    }
}

struct CountingSink {
    cap: usize,
    bytes: usize,
    writes: usize,
    digest: Sha256,
}

impl CountingSink {
    fn new(cap: usize) -> Self {
        Self {
            cap,
            bytes: 0,
            writes: 0,
            digest: Sha256::new(),
        }
    }
}

impl Write for CountingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.bytes.saturating_add(bytes.len()) > self.cap {
            return Err(io::Error::new(io::ErrorKind::WriteZero, "bounded sink cap"));
        }
        self.writes = self.writes.saturating_add(1);
        self.bytes = self.bytes.saturating_add(bytes.len());
        self.digest.update(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// Adapt the shared codecs which expose `fmt::Write` to the bounded sink
/// without first collecting their complete output in a `String`.
struct FormattingSink<'a, W: Write> {
    writer: &'a mut W,
    attempted: usize,
    error: Option<io::Error>,
}

impl<'a, W: Write> FormattingSink<'a, W> {
    fn new(writer: &'a mut W) -> Self {
        Self {
            writer,
            attempted: 0,
            error: None,
        }
    }

    fn failure(&mut self, fallback: impl Into<String>) -> ProbeFailure {
        let message = fallback.into();
        match self.error.take() {
            Some(error) => io_failure(error, self.attempted),
            None => invalid_input(message, self.attempted),
        }
    }

    fn write_text(&mut self, value: &str) -> Result<(), ProbeFailure> {
        if fmt::Write::write_str(self, value).is_err() {
            Err(self.failure("formatted sink write failed"))
        } else {
            Ok(())
        }
    }
}

impl<W: Write> fmt::Write for FormattingSink<'_, W> {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        self.attempted = self.attempted.saturating_add(value.len());
        match self.writer.write_all(value.as_bytes()) {
            Ok(()) => Ok(()),
            Err(error) => {
                self.error = Some(error);
                Err(fmt::Error)
            },
        }
    }
}

fn write_text_body_to<W: Write>(bytes: &[u8], writer: &mut W) -> Result<(), ProbeFailure> {
    let body = dml::text::body::reader::read_with_options(
        bytes,
        quick_xml::name::NamespaceResolver::default(),
        dml::text::body::reader::Limits::DEFAULT,
    )
    .map_err(|error| classify(error, bytes.len()))?;
    let mut sink = FormattingSink::new(writer);
    sink.write_text(
        "<a:txBody xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">",
    )?;
    if dml::text::body::writer::write(&mut sink, &body).is_err() {
        return Err(sink.failure("text body writer failed"));
    }
    sink.write_text("</a:txBody>")
}

fn write_geometry_to<W: Write>(bytes: &[u8], writer: &mut W) -> Result<(), ProbeFailure> {
    let geometry = parse_geometry(bytes)?;
    let mut sink = FormattingSink::new(writer);
    dml::geometry::writer::write_custom_geometry(&mut sink, &geometry);
    if sink.error.is_some() {
        return Err(sink.failure("geometry writer failed"));
    }
    Ok(())
}

fn write_blip_to<W: Write>(bytes: &[u8], writer: &mut W) -> Result<(), ProbeFailure> {
    let relationship = dml::blip::find_first_embed(bytes)
        .map_err(|error| classify(error, bytes.len()))?
        .ok_or_else(|| invalid_input("image blip has no embedded relationship", bytes.len()))?;
    let mut sink = FormattingSink::new(writer);
    if dml::blip::write_embed(&mut sink, &relationship, true).is_err() {
        return Err(sink.failure("image blip writer failed"));
    }
    Ok(())
}

fn write_receipt(sink: CountingSink) -> SinkReceipt {
    SinkReceipt {
        writes: sink.writes,
        bytes: sink.bytes,
        sha256: format!("{:x}", sink.digest.finalize()),
    }
}

fn fixture_seed(fixture: &FixtureLane) -> String {
    let mut seed = fixture.root.clone();
    seed.truncate(seed.len().min(32));
    seed.chars()
        .filter(|character| character.is_ascii_alphanumeric())
        .collect()
}

fn fixture_lane_identity(lane: &str, fixture: &FixtureLane) -> String {
    sha256(
        format!(
            "drawingml-fixture-v1\0{lane}\0{}\0{}\0{}",
            fixture.root, fixture.mutation, fixture.malformed
        )
        .as_bytes(),
    )
}

fn fixture_part_identity(
    lane_identity: &str,
    part_id: &str,
    source_sha256: &str,
    changed_sha256: &str,
    malformed_sha256: &str,
) -> String {
    sha256(
        format!(
            "drawingml-fixture-part-v1\0{lane_identity}\0{part_id}\0{source_sha256}\0{changed_sha256}\0{malformed_sha256}"
        )
        .as_bytes(),
    )
}

fn malformed_fragment(part_id: &str, seed: &str) -> Vec<u8> {
    format!("{seed}<!--drawingml-part:{part_id}-->").into_bytes()
}

fn chart_parts(fixture: &FixtureLane) -> Result<Vec<Part>, ProbeFailure> {
    use dml::chart::authoring::{Kind, Series, Spec, Variant};
    let seed = fixture_seed(fixture);
    let spec = Spec::new(Variant::new(Kind::Column))
        .with_title(format!("DrawingML {seed}"))
        .with_categories(["Q1", "Q2", "Q3"])
        .with_series(Series::new("East", vec![1.0, 2.0, 3.0]));
    let chart = spec.to_chart().map_err(|error| classify(error, 0))?;
    let mut source_wire = Vec::new();
    dml::chart::writer::write(&mut source_wire, &chart).map_err(|error| io_failure(error, 0))?;
    let source = read_chart(&source_wire)?;
    let changed_spec = Spec::new(Variant::new(Kind::Column))
        .with_title(format!("{} changed", fixture.mutation))
        .with_categories(["Q1", "Q2", "Q3"])
        .with_series(Series::new("East", vec![1.0, 2.0, 4.0]));
    let changed_chart = changed_spec
        .to_chart()
        .map_err(|error| classify(error, 0))?;
    let mut changed_wire = Vec::new();
    dml::chart::writer::write(&mut changed_wire, &changed_chart)
        .map_err(|error| io_failure(error, 0))?;
    let changed = read_chart(&changed_wire)?;
    Ok(vec![Part {
        id: "chart",
        kind: PartKind::Chart,
        source,
        changed,
        malformed: malformed_fragment("chart", &fixture.malformed),
        over: Vec::new(),
        limit_resource: "chart editor XML",
    }])
}

fn format_code_part(fixture: &FixtureLane) -> Result<Part, ProbeFailure> {
    let mut value = dml::chart::extension::formatcode2::Element::new("0.00")
        .map_err(|error| classify(error, 0))?;
    let source =
        dml::chart::extension::formatcode2::write(&value).map_err(|error| classify(error, 0))?;
    value
        .set_value(format!("0.00;{}", fixture.mutation.len()))
        .map_err(|error| classify(error, 0))?;
    let changed =
        dml::chart::extension::formatcode2::write(&value).map_err(|error| classify(error, 0))?;
    Ok(Part {
        id: "formatcode2",
        kind: PartKind::FormatCode,
        source,
        changed,
        malformed: malformed_fragment("formatcode2", &fixture.malformed),
        over: vec![b' '; dml::chart::extension::formatcode2::MAX_XML_BYTES + 1],
        limit_resource: "formatcode2 XML bytes",
    })
}

fn diagram_parts(fixture: &FixtureLane) -> Result<Vec<Part>, ProbeFailure> {
    use dml::diagram::{DiagramType, SmartArtBuilder, generate_smartart_data_xml};
    let diagram = SmartArtBuilder::new(DiagramType::Relationship)
        .layout_name("relationship")
        .add_items([fixture_seed(fixture), "second"])
        .build();
    let source = generate_smartart_data_xml(&diagram).into_bytes();
    let outline = dml::diagram::edit::Outline::read(&source)
        .map_err(|error| classify(error, source.len()))?;
    let entry = outline
        .entries
        .first()
        .ok_or_else(|| fail("diagram authoring produced no content entry"))?;
    let text_changed =
        dml::diagram::edit::set_text(&source, &entry.id, &format!("{} changed", fixture.mutation))
            .map_err(|error| classify(error, source.len()))?;
    let changed = dml::diagram::edit::set_drawing_reference(&text_changed, Some("rIdDrawing"))
        .map_err(|error| classify(error, text_changed.len()))?;
    Ok(vec![Part {
        id: "diagram-data-model",
        kind: PartKind::Diagram,
        source,
        changed,
        malformed: malformed_fragment("diagram-data-model", &fixture.malformed),
        over: vec![b' '; dml::diagram::edit::MAX_PART_BYTES + 1],
        limit_resource: "diagram data-model bytes",
    }])
}

fn text_source() -> Result<Vec<u8>, ProbeFailure> {
    use dml::text::body::{Body, Paragraph, Run};
    let body = Body {
        paragraphs: vec![Paragraph {
            runs: vec![Run::plain("seed")],
            ..Paragraph::default()
        }],
        ..Body::default()
    };
    let mut content = String::new();
    dml::text::body::writer::write(&mut content, &body).map_err(|error| classify(error, 0))?;
    Ok(format!(
        "<a:txBody xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">{content}</a:txBody>"
    )
    .into_bytes())
}

fn text_media_parts(fixture: &FixtureLane) -> Result<Vec<Part>, ProbeFailure> {
    let source = text_source()?;
    let mut body =
        dml::text::body::reader::read(&source).map_err(|error| classify(error, source.len()))?;
    body.paragraphs[0].runs[0].text = format!("{} changed", fixture.mutation);
    let mut content = String::new();
    dml::text::body::writer::write(&mut content, &body)
        .map_err(|error| classify(error, source.len()))?;
    let changed = format!(
        "<a:txBody xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">{content}</a:txBody>"
    )
    .into_bytes();
    let svg_wire = br#"<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdSvg"/>"#.to_vec();
    let svg_source = read_svg(&svg_wire)?;
    let mut svg =
        dml::svg_blip::read(&svg_source).map_err(|error| classify(error, svg_source.len()))?;
    svg.set_reference(
        dml::svg_blip::Reference::embedded("rIdSvgChanged")
            .map_err(|error| fail(error.to_string()))?,
    )
    .map_err(|error| classify(error, svg_source.len()))?;
    let svg_changed =
        dml::svg_blip::write(&svg).map_err(|error| classify(error, svg_source.len()))?;
    let model_wire = br#"<m3d:model3d xmlns:m3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdModel" r:link="rIdLinked"><m3d:spPr/><m3d:camera/><m3d:trans/><m3d:raster rName="Office Renderer" rVer="17.0"><m3d:blip r:embed="rIdRaster"/></m3d:raster><m3d:objViewport/></m3d:model3d>"#.to_vec();
    let model_source = read_model3d(&model_wire)?;
    let mut metadata = dml::model3d::codec::read(&model_source)
        .map_err(|error| classify(error, model_source.len()))?;
    metadata.reference = dml::model3d::Reference::embedded("rIdModelChanged")
        .map_err(|error| fail(error.to_string()))?;
    let model_changed = dml::model3d::codec::write(&metadata)
        .map_err(|error| classify(error, model_source.len()))?;
    let mut blip = String::new();
    dml::blip::write_embed(&mut blip, "rIdImage", true).map_err(|error| fail(error.to_string()))?;
    let blip_source = blip.into_bytes();
    let blip_changed = br#"<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdImageChanged"/>"#.to_vec();
    Ok(vec![
        Part {
            id: "text-body",
            kind: PartKind::TextBody,
            source,
            changed,
            malformed: malformed_fragment("text-body", &fixture.malformed),
            over: vec![b' '; 64 * 1024 * 1024 + 1],
            limit_resource: "text body input bytes",
        },
        Part {
            id: "svg-blip",
            kind: PartKind::SvgBlip,
            source: svg_source,
            changed: svg_changed,
            malformed: malformed_fragment("svg-blip", &fixture.malformed),
            over: vec![b' '; dml::svg_blip::MAX_XML_BYTES + 1],
            limit_resource: "SVG blip XML bytes",
        },
        Part {
            id: "model3d",
            kind: PartKind::Model3d,
            source: model_source,
            changed: model_changed,
            malformed: malformed_fragment("model3d", &fixture.malformed),
            over: vec![b' '; 16 * 1024 * 1024 + 1],
            limit_resource: "model3d XML bytes",
        },
        Part {
            id: "image-blip",
            kind: PartKind::Blip,
            source: blip_source,
            changed: blip_changed,
            malformed: malformed_fragment("image-blip", &fixture.malformed),
            over: Vec::new(),
            limit_resource: "blip XML",
        },
    ])
}

fn theme_parts(fixture: &FixtureLane) -> Result<Vec<Part>, ProbeFailure> {
    use dml::theme::{Color, Face, FontSet, Palette, Slot};
    let mut palette = Palette::new("Office");
    for slot in Slot::ALL {
        palette = palette.with(
            slot,
            Color::rgb("336699").map_err(|error| classify(error, 0))?,
        );
    }
    let fonts = FontSet::new(
        "Office",
        Face::new("Aptos")
            .east_asian("Aptos")
            .complex_script("Aptos"),
        Face::new("Aptos")
            .east_asian("Aptos")
            .complex_script("Aptos"),
    );
    let source = dml::theme::codec::encode_part("Office", &palette, &fonts)
        .map_err(|error| classify(error, 0))?;
    let changed = dml::theme::codec::encode_part("Office Changed", &palette, &fonts)
        .map_err(|error| classify(error, 0))?;
    let family = dml::theme::family::Family::new(
        "Office",
        "{11111111-1111-1111-1111-111111111111}",
        "{22222222-2222-2222-2222-222222222222}",
    )
    .map_err(|error| classify(error, 0))?;
    let family_source = dml::theme::family::write(&family).map_err(|error| classify(error, 0))?;
    let family_changed = family
        .with_name(format!("Family {}", fixture.mutation.len()))
        .map_err(|error| classify(error, family_source.len()))?;
    let family_changed = dml::theme::family::write(&family_changed)
        .map_err(|error| classify(error, family_source.len()))?;
    let family_part_source = complete_theme_family_part(&family_source)?;
    let family_part_changed = complete_theme_family_part(&family_changed)?;
    let _ = dml::theme::family::part::read(&family_part_source)
        .map_err(|error| classify(error, family_part_source.len()))?;
    let ink_source = br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" lengthUnit="cm" timeUnit="s"><iact:action type="add" startTime="0.25"><iact:actionData name="stroke"/></iact:action><iact:actionGroup type="transform" startTime="1"><iact:action type="remove" startTime="2"/></iact:actionGroup></iact:actions>"#.to_vec();
    let ink_profile = dml::ink::actions::read_profile(&ink_source)
        .map_err(|error| classify(error, ink_source.len()))?;
    let ink_changed = ink_source
        .windows(4)
        .position(|window| window == b"0.25")
        .map(|position| {
            let mut output = ink_source.clone();
            output.splice(position..position + 4, b"0.50".iter().copied());
            output
        })
        .unwrap_or_else(|| ink_source.clone());
    let _ = dml::ink::actions::read_profile(&ink_changed)
        .map_err(|error| classify(error, ink_source.len()))?;
    let _ = dml::ink::actions::write_profile(&ink_profile)
        .map_err(|error| classify(error, ink_source.len()))?;
    Ok(vec![
        Part {
            id: "theme",
            kind: PartKind::Theme,
            source,
            changed,
            malformed: malformed_fragment("theme", &fixture.malformed),
            over: vec![b' '; dml::theme::codec::MAX_XML_BYTES + 1],
            limit_resource: "theme XML bytes",
        },
        Part {
            id: "theme-family",
            kind: PartKind::ThemeFamily,
            source: family_source,
            changed: family_changed,
            malformed: malformed_fragment("theme-family", &fixture.malformed),
            over: vec![b' '; dml::theme::family::MAX_XML_BYTES + 1],
            limit_resource: "theme family XML bytes",
        },
        Part {
            id: "theme-family-part",
            kind: PartKind::ThemeFamilyPart,
            source: family_part_source,
            changed: family_part_changed,
            malformed: malformed_fragment("theme-family-part", &fixture.malformed),
            over: vec![b' '; dml::theme::family::part::MAX_XML_BYTES + 1],
            limit_resource: "Theme Family XML bytes",
        },
        Part {
            id: "ink-actions",
            kind: PartKind::InkActions,
            source: ink_source,
            changed: ink_changed,
            malformed: malformed_fragment("ink-actions", &fixture.malformed),
            over: vec![b' '; dml::ink::MAX_SOURCE_BYTES + 1],
            limit_resource: "ink actions source bytes",
        },
    ])
}

fn geometry_parts(fixture: &FixtureLane) -> Result<Vec<Part>, ProbeFailure> {
    use dml::geometry::{AdjustValue, CustomGeometry, Path, PathCommand, Point};
    let geometry = CustomGeometry::new().with_path(
        Path::new(100, 100)
            .with_command(PathCommand::MoveTo(Point::new(0_i64, 0_i64)))
            .with_command(PathCommand::LineTo(Point::new(100_i64, 100_i64)))
            .with_command(PathCommand::Close),
    );
    let mut geometry_xml = String::new();
    dml::geometry::writer::write_custom_geometry(&mut geometry_xml, &geometry);
    let source = geometry_xml.into_bytes();
    let mut changed_geometry = geometry.clone();
    if let Some(PathCommand::LineTo(point)) = changed_geometry.paths[0].commands.get_mut(1) {
        point.x = AdjustValue::Value(90);
    }
    let mut changed_geometry_xml = String::new();
    dml::geometry::writer::write_custom_geometry(&mut changed_geometry_xml, &changed_geometry);
    let color_source =
        dml::color::write(&dml::color::Value::rgb(dml::color::Rgb::new(51, 102, 153)))
            .map_err(|error| classify(error, 0))?;
    let color_changed =
        dml::color::write(&dml::color::Value::rgb(dml::color::Rgb::new(153, 102, 51)))
            .map_err(|error| classify(error, 0))?;
    let format_part = format_code_part(fixture)?;
    let transform = dml::transform::Transform::new()
        .with_offset(dml::transform::Point::new(
            dml::coordinate::Coordinate::from(0),
            dml::coordinate::Coordinate::from(0),
        ))
        .with_extent(dml::transform::Size::new(
            dml::coordinate::Extent::from(100_u32),
            dml::coordinate::Extent::from(100_u32),
        ));
    let transform_source =
        dml::transform::codec::write(&transform).map_err(|error| classify(error, 0))?;
    let transform_changed =
        dml::transform::codec::write(&transform.with_rotation(dml::transform::Angle::new(60000)))
            .map_err(|error| classify(error, 0))?;
    Ok(vec![
        Part {
            id: "geometry",
            kind: PartKind::Geometry,
            source,
            changed: changed_geometry_xml.into_bytes(),
            malformed: malformed_fragment("geometry", &fixture.malformed),
            over: Vec::new(),
            limit_resource: "geometry structural limits",
        },
        Part {
            id: "color",
            kind: PartKind::Color,
            source: color_source,
            changed: color_changed,
            malformed: malformed_fragment("color", &fixture.malformed),
            over: vec![b' '; dml::color::MAX_XML_BYTES + 1],
            limit_resource: "DrawingML color XML",
        },
        format_part,
        Part {
            id: "transform",
            kind: PartKind::Transform,
            source: transform_source,
            changed: transform_changed,
            malformed: malformed_fragment("transform", &fixture.malformed),
            over: vec![b' '; (1 << 20) + 1],
            limit_resource: "DrawingML transform XML bytes",
        },
    ])
}

fn owner_fixture(route: OwnerRoute, parts: &[Part]) -> Result<OwnerFixture, ProbeFailure> {
    let kind = match route {
        OwnerRoute::DrawingmlChart => PartKind::Chart,
        OwnerRoute::DrawingmlThemeFamily => PartKind::ThemeFamily,
        OwnerRoute::DrawingmlInkActions => PartKind::InkActions,
        OwnerRoute::DrawingmlTransform => PartKind::Transform,
    };
    let part = parts
        .iter()
        .find(|part| std::mem::discriminant(&part.kind) == std::mem::discriminant(&kind))
        .ok_or_else(|| fail(format!("missing {} owner fixture", owner_route_name(route))))?;
    Ok(OwnerFixture {
        route,
        source: part.source.clone(),
        // Use a second valid semantic state as the stale recipient.  The
        // production patch precondition, rather than a local byte-copy
        // helper, must reject it.
        stale: part.changed.clone(),
    })
}

fn make_case(lane: &str, fixture: &FixtureLane) -> Result<LaneCase, ProbeFailure> {
    let fixture_identity_sha256 = fixture_lane_identity(lane, fixture);
    let parts = match lane {
        "drawingml-chart-raw" => chart_parts(fixture)?,
        "drawingml-diagram-raw" => diagram_parts(fixture)?,
        "drawingml-text-media-raw" => text_media_parts(fixture)?,
        "drawingml-theme-ink" => theme_parts(fixture)?,
        "drawingml-geometry-color-formatcode2-transform" => geometry_parts(fixture)?,
        _ => return Err(fail(format!("unknown lane {lane}"))),
    };
    let mut stale = parts
        .first()
        .map(|part| part.source.clone())
        .unwrap_or_default();
    stale.extend_from_slice(b"stale-source");
    let owners = owner_routes(lane)
        .iter()
        .copied()
        .map(|route| owner_fixture(route, &parts))
        .collect::<Result<Vec<_>, _>>()?;
    Ok(LaneCase {
        parts,
        stale,
        owners,
        fixture_identity_sha256,
    })
}

fn select_owner_route(case: &mut LaneCase, route: &str) -> Result<(), ProbeFailure> {
    let route = owner_route_from_name(route)
        .ok_or_else(|| fail(format!("unknown DrawingML owner route {route:?}")))?;
    case.owners.retain(|fixture| fixture.route == route);
    if case.owners.len() != 1 {
        return Err(fail(format!(
            "owner route {} is not available in lane",
            owner_route_name(route)
        )));
    }
    Ok(())
}

fn select_raw_scope(case: &mut LaneCase, scope: &str) -> Result<(), ProbeFailure> {
    match scope {
        "bulk" => case
            .parts
            .retain(|part| matches!(part.kind, PartKind::Theme | PartKind::InkActions)),
        "family" => case
            .parts
            .retain(|part| matches!(part.kind, PartKind::ThemeFamilyPart)),
        _ => return Err(fail(format!("unknown DrawingML raw scope {scope:?}"))),
    }
    if case.parts.is_empty() {
        return Err(fail(format!("raw scope {scope} has no fixture parts")));
    }
    case.owners.clear();
    Ok(())
}

#[cfg(not(feature = "drawingml-owner-current"))]
fn owner_error(route: OwnerRoute, error: impl std::fmt::Display) -> ProbeFailure {
    fail(format!("{}: {error}", owner_route_name(route)))
}

#[cfg(feature = "drawingml-owner-current")]
fn owner_error(route: OwnerRoute, error: dml::Error) -> ProbeFailure {
    let message = error.to_string();
    let refusal = if matches!(&error, dml::Error::PatchConflict) {
        Some(RefusalReceipt {
            kind: "patch_conflict",
            source_kind: Some("patch_conflict"),
            limit_kind: None,
            resource: None,
            limit: None,
            observed: 0,
            message: message.clone(),
        })
    } else {
        None
    };
    ProbeFailure { message, refusal }
}

#[cfg(not(feature = "drawingml-owner-current"))]
fn owner_commit(
    fixture: &OwnerFixture,
    changed: bool,
) -> Result<(Vec<u8>, OwnerPatch), ProbeFailure> {
    match fixture.route {
        OwnerRoute::DrawingmlChart => {
            let snapshot = dml::chart::edit::Snapshot::from_xml(fixture.source.clone())
                .map_err(|error| owner_error(fixture.route, error))?;
            let mut edit = snapshot.edit();
            if changed {
                edit.set_title(Some(dml::chart::data::TitleText::from_string(
                    "DrawingML qualification mutation",
                )))
                .map_err(|error| owner_error(fixture.route, error))?;
            }
            let commit = edit
                .commit()
                .map_err(|error| owner_error(fixture.route, error))?;
            Ok((
                commit.snapshot().xml_bytes().to_vec(),
                OwnerPatch::DrawingmlChart(commit.patch().clone()),
            ))
        },
        OwnerRoute::DrawingmlThemeFamily => {
            let snapshot = dml::theme::family::Snapshot::from_xml(fixture.source.clone())
                .map_err(|error| owner_error(fixture.route, error))?;
            let mut edit = snapshot.edit();
            if changed {
                edit.set_name("DrawingML qualification mutation")
                    .map_err(|error| owner_error(fixture.route, error))?;
            }
            let commit = edit
                .commit()
                .map_err(|error| owner_error(fixture.route, error))?;
            Ok((
                commit.snapshot().xml_bytes().to_vec(),
                OwnerPatch::DrawingmlThemeFamily(commit.patch().clone()),
            ))
        },
        OwnerRoute::DrawingmlInkActions => {
            use dml::ink::actions::{ActionSelector, Edit};
            let profile = dml::ink::actions::read_profile(&fixture.source)
                .map_err(|error| owner_error(fixture.route, error))?;
            let mut edit =
                Edit::from_profile(&profile).map_err(|error| owner_error(fixture.route, error))?;
            if changed {
                edit.set_start_time(ActionSelector::ordinal(0), "0.50")
                    .map_err(|error| owner_error(fixture.route, error))?;
            }
            let commit = edit
                .commit()
                .map_err(|error| owner_error(fixture.route, error))?;
            Ok((
                commit
                    .to_vec()
                    .map_err(|error| owner_error(fixture.route, error))?,
                OwnerPatch::DrawingmlInkActions(commit.patch().clone()),
            ))
        },
        OwnerRoute::DrawingmlTransform => {
            let snapshot = dml::transform::Snapshot::from_xml(fixture.source.clone())
                .map_err(|error| owner_error(fixture.route, error))?;
            let mut edit = snapshot.edit();
            if changed {
                edit.set_rotation(Some(dml::transform::Angle::new(60000)));
            }
            let commit = edit
                .commit()
                .map_err(|error| owner_error(fixture.route, error))?;
            Ok((
                commit.snapshot().xml_bytes().to_vec(),
                OwnerPatch::DrawingmlTransform(commit.patch().clone()),
            ))
        },
    }
}

#[cfg(feature = "drawingml-owner-current")]
fn owner_commit(
    fixture: &OwnerFixture,
    changed: bool,
) -> Result<(Vec<u8>, OwnerPatch), ProbeFailure> {
    match fixture.route {
        OwnerRoute::DrawingmlChart => {
            let package = dml::chart::edit::Package::from_xml(fixture.source.clone())
                .map_err(|error| owner_error(fixture.route, error))?;
            let mut edit = package.edit();
            if changed {
                edit.set_title(Some(dml::chart::data::TitleText::from_string(
                    "DrawingML qualification mutation",
                )))
                .map_err(|error| owner_error(fixture.route, error))?;
            }
            let commit = edit
                .commit()
                .map_err(|error| owner_error(fixture.route, error))?;
            Ok((
                commit.package().xml_bytes().to_vec(),
                OwnerPatch::DrawingmlChart(commit.patch().clone()),
            ))
        },
        OwnerRoute::DrawingmlThemeFamily => {
            let package = dml::theme::family::Package::from_xml(fixture.source.clone())
                .map_err(|error| owner_error(fixture.route, error))?;
            let mut edit = package.edit();
            if changed {
                edit.set_name("DrawingML qualification mutation")
                    .map_err(|error| owner_error(fixture.route, error))?;
            }
            let commit = edit
                .commit()
                .map_err(|error| owner_error(fixture.route, error))?;
            Ok((
                commit.package().xml_bytes().to_vec(),
                OwnerPatch::DrawingmlThemeFamily(commit.patch().clone()),
            ))
        },
        OwnerRoute::DrawingmlInkActions => {
            use dml::ink::actions::{ActionSelector, Edit, Limits};
            let profile = dml::ink::actions::read_profile(&fixture.source)
                .map_err(|error| owner_error(fixture.route, error))?;
            let mut edit = Edit::from_profile(&profile, Limits::default())
                .map_err(|error| owner_error(fixture.route, error))?;
            if changed {
                edit.set_start_time(ActionSelector::ordinal(0), "0.50")
                    .map_err(|error| owner_error(fixture.route, error))?;
            }
            let commit = edit
                .commit()
                .map_err(|error| owner_error(fixture.route, error))?;
            Ok((
                commit.package().as_bytes().to_vec(),
                OwnerPatch::DrawingmlInkActions(commit.patch().clone()),
            ))
        },
        OwnerRoute::DrawingmlTransform => {
            let package = dml::transform::Package::from_xml(fixture.source.clone())
                .map_err(|error| owner_error(fixture.route, error))?;
            let mut edit = package.edit();
            if changed {
                edit.set_rotation(Some(dml::transform::Angle::new(60000)));
            }
            let commit = edit
                .commit()
                .map_err(|error| owner_error(fixture.route, error))?;
            Ok((
                commit.package().xml_bytes().to_vec(),
                OwnerPatch::DrawingmlTransform(commit.patch().clone()),
            ))
        },
    }
}

#[cfg(not(feature = "drawingml-owner-current"))]
fn owner_apply(
    route: OwnerRoute,
    source: &[u8],
    patch: &OwnerPatch,
    inverse: bool,
) -> Result<Vec<u8>, ProbeFailure> {
    match (route, patch) {
        (OwnerRoute::DrawingmlChart, OwnerPatch::DrawingmlChart(patch)) => {
            let snapshot = dml::chart::edit::Snapshot::from_xml(source.to_vec())
                .map_err(|error| owner_error(route, error))?;
            let patch = if inverse {
                patch.clone().inverse()
            } else {
                patch.clone()
            };
            patch
                .apply(&snapshot)
                .map(|snapshot| snapshot.xml_bytes().to_vec())
                .map_err(|error| owner_error(route, error))
        },
        (OwnerRoute::DrawingmlThemeFamily, OwnerPatch::DrawingmlThemeFamily(patch)) => {
            let snapshot = dml::theme::family::Snapshot::from_xml(source.to_vec())
                .map_err(|error| owner_error(route, error))?;
            let patch = if inverse {
                patch.clone().inverse()
            } else {
                patch.clone()
            };
            patch
                .apply(&snapshot)
                .map(|snapshot| snapshot.xml_bytes().to_vec())
                .map_err(|error| owner_error(route, error))
        },
        (OwnerRoute::DrawingmlInkActions, OwnerPatch::DrawingmlInkActions(patch)) => {
            let patch = if inverse {
                patch.inverse()
            } else {
                patch.clone()
            };
            patch
                .apply(source)
                .map_err(|error| owner_error(route, error))
        },
        (OwnerRoute::DrawingmlTransform, OwnerPatch::DrawingmlTransform(patch)) => {
            let snapshot = dml::transform::Snapshot::from_xml(source.to_vec())
                .map_err(|error| owner_error(route, error))?;
            let patch = if inverse {
                patch.clone().inverse()
            } else {
                patch.clone()
            };
            patch
                .apply(&snapshot)
                .map(|snapshot| snapshot.xml_bytes().to_vec())
                .map_err(|error| owner_error(route, error))
        },
        _ => Err(fail("owner route and patch type differ")),
    }
}

#[cfg(feature = "drawingml-owner-current")]
fn owner_apply(
    route: OwnerRoute,
    source: &[u8],
    patch: &OwnerPatch,
    inverse: bool,
) -> Result<Vec<u8>, ProbeFailure> {
    match (route, patch) {
        (OwnerRoute::DrawingmlChart, OwnerPatch::DrawingmlChart(patch)) => {
            let package = dml::chart::edit::Package::from_xml(source.to_vec())
                .map_err(|error| owner_error(route, error))?;
            let patch = if inverse {
                patch.clone().inverse()
            } else {
                patch.clone()
            };
            package
                .apply(&patch)
                .map(|commit| commit.package().xml_bytes().to_vec())
                .map_err(|error| owner_error(route, error))
        },
        (OwnerRoute::DrawingmlThemeFamily, OwnerPatch::DrawingmlThemeFamily(patch)) => {
            let package = dml::theme::family::Package::from_xml(source.to_vec())
                .map_err(|error| owner_error(route, error))?;
            let patch = if inverse {
                patch.clone().inverse()
            } else {
                patch.clone()
            };
            package
                .apply(&patch)
                .map(|commit| commit.package().xml_bytes().to_vec())
                .map_err(|error| owner_error(route, error))
        },
        (OwnerRoute::DrawingmlInkActions, OwnerPatch::DrawingmlInkActions(patch)) => {
            let patch = if inverse {
                patch.inverse()
            } else {
                patch.clone()
            };
            patch
                .apply(source)
                .map_err(|error| owner_error(route, error))
        },
        (OwnerRoute::DrawingmlTransform, OwnerPatch::DrawingmlTransform(patch)) => {
            let package = dml::transform::Package::from_xml(source.to_vec())
                .map_err(|error| owner_error(route, error))?;
            let patch = if inverse {
                patch.clone().inverse()
            } else {
                patch.clone()
            };
            package
                .apply(&patch)
                .map(|commit| commit.package().xml_bytes().to_vec())
                .map_err(|error| owner_error(route, error))
        },
        _ => Err(fail("owner route and patch type differ")),
    }
}

fn owner_replay(
    fixture: &OwnerFixture,
    patch: &OwnerPatch,
    inverse: bool,
) -> Result<Vec<u8>, ProbeFailure> {
    let (source, inverse) = if inverse {
        (owner_commit(fixture, true)?.0, true)
    } else {
        (fixture.source.clone(), false)
    };
    owner_apply(fixture.route, &source, patch, inverse)
}

fn owner_stale_refusal(
    fixture: &OwnerFixture,
    patch: &OwnerPatch,
) -> Result<RefusalReceipt, ProbeFailure> {
    match owner_apply(fixture.route, &fixture.stale, patch, false) {
        Ok(_) => Err(fail(format!(
            "{} accepted a stale patch",
            owner_route_name(fixture.route)
        ))),
        Err(error) => {
            #[cfg(feature = "drawingml-owner-current")]
            {
                let Some(mut refusal) = error.refusal else {
                    return Err(fail(format!(
                        "{} stale replay returned an untyped error: {}",
                        owner_route_name(fixture.route),
                        error.message
                    )));
                };
                if refusal.kind != "patch_conflict" || refusal.source_kind != Some("patch_conflict")
                {
                    return Err(fail(format!(
                        "{} stale replay returned the wrong typed error: {}",
                        owner_route_name(fixture.route),
                        refusal.message
                    )));
                }
                refusal.observed = fixture.stale.len();
                Ok(refusal)
            }
            #[cfg(not(feature = "drawingml-owner-current"))]
            {
                // The pre-owner baseline rejects stale patches as its generic
                // invalid-input error.  Keep that historical behavior in the
                // receipt while naming the normalized contract kind used by
                // the paired validator.
                Ok(RefusalReceipt {
                    kind: "patch_conflict",
                    source_kind: Some("invalid_input"),
                    limit_kind: None,
                    resource: None,
                    limit: None,
                    observed: fixture.stale.len(),
                    message: error.message,
                })
            }
        },
    }
}

fn owner_proof(
    fixture_identity_sha256: &str,
    fixture: &OwnerFixture,
) -> Result<OwnerCommitReceipt, ProbeFailure> {
    let (noop_bytes, _noop_patch) = owner_commit(fixture, false)?;
    if noop_bytes != fixture.source {
        return Err(fail(format!(
            "{} no-op commit changed the exact source bytes",
            owner_route_name(fixture.route)
        )));
    }
    let (changed_bytes, patch) = owner_commit(fixture, true)?;
    let forward_bytes = owner_replay(fixture, &patch, false)?;
    let inverse_bytes = owner_replay(fixture, &patch, true)?;
    if forward_bytes != changed_bytes {
        return Err(fail(format!(
            "{} forward replay differs from committed output",
            owner_route_name(fixture.route)
        )));
    }
    if inverse_bytes != fixture.source {
        return Err(fail(format!(
            "{} inverse replay did not restore the exact source",
            owner_route_name(fixture.route)
        )));
    }
    let stale_refusal = owner_stale_refusal(fixture, &patch)?;
    Ok(OwnerCommitReceipt {
        route: owner_route_name(fixture.route),
        fixture_identity_sha256: sha256(
            format!(
                "drawingml-owner-fixture-v1\0{fixture_identity_sha256}\0{}\0{}",
                owner_route_name(fixture.route),
                sha256(&fixture.source),
                sha256(&fixture.stale),
            )
            .as_bytes(),
        ),
        source_sha256: sha256(&fixture.source),
        noop_sha256: sha256(&noop_bytes),
        stale_sha256: sha256(&fixture.stale),
        changed_sha256: sha256(&changed_bytes),
        forward_sha256: sha256(&forward_bytes),
        inverse_sha256: sha256(&inverse_bytes),
        stale_kind: stale_refusal.kind,
        stale_refusal,
    })
}

fn raw_fidelity(case: &LaneCase) -> Result<Vec<RawFidelityReceipt>, ProbeFailure> {
    case.parts
        .iter()
        .map(|part| {
            let source_output = roundtrip(part.kind, &part.source)?;
            let changed_output = roundtrip(part.kind, &part.changed)?;
            if source_output.is_empty() || changed_output.is_empty() {
                return Err(fail(format!("{} raw readback was empty", part.id)));
            }
            if source_output == changed_output {
                return Err(fail(format!(
                    "{} changed fixture did not change its raw readback",
                    part.id
                )));
            }
            let source_sha256 = sha256(&part.source);
            let changed_sha256 = sha256(&part.changed);
            let malformed_sha256 = sha256(&part.malformed);
            Ok(RawFidelityReceipt {
                part_id: part.id,
                fixture_identity_sha256: fixture_part_identity(
                    &case.fixture_identity_sha256,
                    part.id,
                    &source_sha256,
                    &changed_sha256,
                    &malformed_sha256,
                ),
                source_sha256,
                source_output_sha256: sha256(&source_output),
                changed_sha256,
                changed_output_sha256: sha256(&changed_output),
                malformed_sha256,
                changed_output_differs: true,
            })
        })
        .collect()
}

fn size_class_receipts(case: &LaneCase) -> Result<Vec<SizeClassReceipt>, ProbeFailure> {
    let mut receipts = Vec::with_capacity(case.parts.len().saturating_mul(5));
    for part in &case.parts {
        let cases = if input_limit(part.kind).is_some() {
            let maximum = part.over.len().saturating_sub(1);
            vec![
                ("small", part.source.clone(), None, true),
                (
                    "near_limit",
                    sized_with_whitespace(&part.source, maximum.saturating_sub(1)),
                    None,
                    true,
                ),
                (
                    "at_limit",
                    sized_with_whitespace(&part.source, maximum),
                    None,
                    true,
                ),
                ("over_limit", part.over.clone(), Some("limit"), true),
                (
                    "malformed",
                    part.malformed.clone(),
                    Some("invalid_input"),
                    true,
                ),
            ]
        } else {
            vec![
                ("small", part.source.clone(), None, true),
                ("near_limit", part.source.clone(), None, true),
                ("at_limit", part.source.clone(), None, true),
                ("over_limit", Vec::new(), None, false),
                (
                    "malformed",
                    part.malformed.clone(),
                    Some("invalid_input"),
                    true,
                ),
            ]
        };
        for (class, bytes, expected_kind, applicable) in cases {
            if !applicable {
                receipts.push(SizeClassReceipt {
                    part_id: part.id,
                    class,
                    bytes: 0,
                    applicable: false,
                    accepted: false,
                    output_sha256: None,
                    refusal: None,
                });
                continue;
            }
            match roundtrip(part.kind, &bytes) {
                Ok(output) => {
                    if expected_kind.is_some() {
                        return Err(fail(format!(
                            "{} {class} endpoint was accepted despite expected refusal",
                            part.id
                        )));
                    }
                    receipts.push(SizeClassReceipt {
                        part_id: part.id,
                        class,
                        bytes: bytes.len(),
                        applicable: true,
                        accepted: true,
                        output_sha256: Some(sha256(&output)),
                        refusal: None,
                    });
                },
                Err(error) => {
                    let refusal = error.refusal.ok_or_else(|| {
                        fail(format!("{} {class} endpoint lacked typed refusal", part.id))
                    })?;
                    if expected_kind != Some(refusal.kind) {
                        return Err(fail(format!(
                            "{} {class} endpoint returned {} instead of {:?}",
                            part.id, refusal.kind, expected_kind
                        )));
                    }
                    receipts.push(SizeClassReceipt {
                        part_id: part.id,
                        class,
                        bytes: bytes.len(),
                        applicable: true,
                        accepted: false,
                        output_sha256: None,
                        refusal: Some(refusal),
                    });
                },
            }
        }
    }
    Ok(receipts)
}

fn sized_with_whitespace(source: &[u8], target: usize) -> Vec<u8> {
    if source.len() >= target {
        return source.to_vec();
    }
    let mut output = Vec::with_capacity(target);
    output.extend_from_slice(source);
    output.resize(target, b' ');
    output
}

fn input_limit(kind: PartKind) -> Option<(&'static str, usize)> {
    match kind {
        PartKind::FormatCode => Some(("formatcode2 XML bytes", 1 << 20)),
        PartKind::Diagram => Some(("diagram data-model bytes", 16 * 1024 * 1024)),
        PartKind::TextBody => Some(("text body input bytes", 64 * 1024 * 1024)),
        PartKind::SvgBlip => Some(("SVG blip XML bytes", 16 * 1024 * 1024)),
        PartKind::Model3d => Some(("model3d XML bytes", 16 * 1024 * 1024)),
        PartKind::Theme => Some(("theme XML bytes", 8 * 1024 * 1024)),
        PartKind::ThemeFamily | PartKind::ThemeFamilyPart => {
            Some(("theme family XML bytes", 1 << 20))
        },
        PartKind::InkActions => Some(("ink actions source bytes", 16 * 1024 * 1024)),
        PartKind::Color => Some(("DrawingML color XML", 64 * 1024)),
        PartKind::Transform => Some(("DrawingML transform XML", 1 << 20)),
        PartKind::Chart | PartKind::Blip | PartKind::Geometry => None,
    }
}

fn part_has_input_limit(kind: PartKind) -> bool {
    input_limit(kind).is_some()
}

fn parse_geometry(bytes: &[u8]) -> Result<dml::geometry::CustomGeometry, ProbeFailure> {
    use dml::geometry::reader::{CustomGeometryBuilder, GeometryElement};
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut builder = CustomGeometryBuilder::new();
    let mut stack = Vec::new();
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| invalid_input(error.to_string(), bytes.len()))?
            .into_owned();
        match event {
            Event::Start(element) => {
                let local = element.local_name().as_ref().to_vec();
                if stack.is_empty() {
                    if local.as_slice() != b"custGeom" {
                        return Err(invalid_input("geometry root is not custGeom", bytes.len()));
                    }
                    stack.push(GeometryElement::CustomGeometry);
                } else {
                    let parent = *stack
                        .last()
                        .ok_or_else(|| invalid_input("geometry stack is empty", bytes.len()))?;
                    let child = builder
                        .open(parent, &local, &element, reader.decoder())
                        .map_err(|error| classify(error, bytes.len()))?;
                    let child = child.ok_or_else(|| {
                        invalid_input(
                            format!(
                                "unsupported geometry child {:?}",
                                String::from_utf8_lossy(&local)
                            ),
                            bytes.len(),
                        )
                    })?;
                    stack.push(child);
                }
            },
            Event::Empty(element) => {
                let local = element.local_name().as_ref().to_vec();
                if stack.is_empty() {
                    return Err(invalid_input(
                        "empty geometry root is unsupported",
                        bytes.len(),
                    ));
                }
                let parent = *stack
                    .last()
                    .ok_or_else(|| invalid_input("geometry stack is empty", bytes.len()))?;
                let child = builder
                    .open(parent, &local, &element, reader.decoder())
                    .map_err(|error| classify(error, bytes.len()))?;
                let child = child.ok_or_else(|| {
                    invalid_input(
                        format!(
                            "unsupported geometry child {:?}",
                            String::from_utf8_lossy(&local)
                        ),
                        bytes.len(),
                    )
                })?;
                builder
                    .close(child)
                    .map_err(|error| classify(error, bytes.len()))?;
            },
            Event::End(_) => {
                let element = stack
                    .pop()
                    .ok_or_else(|| invalid_input("geometry close underflow", bytes.len()))?;
                builder
                    .close(element)
                    .map_err(|error| classify(error, bytes.len()))?;
                if stack.is_empty() {
                    break;
                }
            },
            Event::Eof => return Err(invalid_input("geometry root is unterminated", bytes.len())),
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_) => {},
            Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(invalid_input(
                    "geometry contains forbidden markup",
                    bytes.len(),
                ));
            },
        }
        buffer.clear();
    }
    builder
        .finish()
        .map_err(|error| classify(error, bytes.len()))
}

fn read_shared_lane(case: &LaneCase) -> Result<Vec<Vec<u8>>, ProbeFailure> {
    case.parts
        .iter()
        .filter(|part| supports_shared(part.kind))
        .map(|part| match part.kind {
            PartKind::FormatCode => {
                let value = dml::chart::extension::formatcode2::read_shared(
                    std::sync::Arc::<[u8]>::from(part.source.clone()),
                )
                .map_err(|error| classify(error, part.source.len()))?;
                dml::chart::extension::formatcode2::write(&value)
                    .map_err(|error| classify(error, part.source.len()))
            },
            PartKind::ThemeFamily => {
                let value = dml::theme::family::read_shared(std::sync::Arc::<[u8]>::from(
                    part.source.clone(),
                ))
                .map_err(|error| classify(error, part.source.len()))?;
                dml::theme::family::write(&value)
                    .map_err(|error| classify(error, part.source.len()))
            },
            PartKind::ThemeFamilyPart => {
                let value = dml::theme::family::part::read_shared(std::sync::Arc::<[u8]>::from(
                    part.source.clone(),
                ))
                .map_err(|error| classify(error, part.source.len()))?;
                Ok(value.xml_bytes().to_vec())
            },
            _ => Err(fail(format!("{} has no shared-reader route", part.id))),
        })
        .collect()
}

fn malformed_check(case: &LaneCase) -> Result<serde_json::Value, ProbeFailure> {
    let mut parts = Vec::with_capacity(case.parts.len());
    for part in &case.parts {
        let error = roundtrip(part.kind, &part.malformed)
            .map(|_| ())
            .expect_err("malformed source must be refused");
        let refusal = error
            .refusal
            .ok_or_else(|| fail(format!("{} malformed input lacked typed refusal", part.id)))?;
        if refusal.kind != "invalid_input" {
            return Err(fail(format!(
                "{} malformed input was not invalid_input",
                part.id
            )));
        }
        parts.push(serde_json::json!({
            "part_id": part.id,
            "refusal": refusal,
        }));
    }
    let limit = limit_check(case)?;
    Ok(serde_json::json!({
        "malformed_refusal": true,
        "refusal": parts.first().and_then(|part| part.get("refusal")).cloned(),
        "parts": parts,
        "limit": limit,
    }))
}

fn limit_check(case: &LaneCase) -> Result<serde_json::Value, ProbeFailure> {
    let mut parts = Vec::new();
    for part in &case.parts {
        if input_limit(part.kind).is_none() {
            continue;
        }
        let error = roundtrip(part.kind, &part.over)
            .map(|_| ())
            .expect_err("over-limit source must be refused");
        let refusal = error
            .refusal
            .ok_or_else(|| fail(format!("{} over-limit input lacked typed refusal", part.id)))?;
        if refusal.kind != "limit"
            || refusal.limit_kind.is_none()
            || refusal.limit.is_none()
            || refusal.observed <= refusal.limit.unwrap_or(0)
        {
            return Err(fail(format!(
                "{} over-limit source was not a typed limit refusal: {}",
                part.id, error.message
            )));
        }
        parts.push(serde_json::json!({
            "part_id": part.id,
            "expected_legacy_resource": part.limit_resource,
            "refusal": refusal,
        }));
    }
    Ok(serde_json::json!({
        "limit_refusal": !parts.is_empty(),
        "parts": parts,
    }))
}

fn theme_family_mutations(
    case: &LaneCase,
) -> Result<Vec<ThemeFamilyMutationReceipt>, ProbeFailure> {
    let part = case
        .parts
        .iter()
        .find(|part| matches!(part.kind, PartKind::ThemeFamilyPart))
        .ok_or_else(|| fail("theme-family mutation route has no complete Theme part"))?;
    let source_snapshot = dml::theme::family::part::read(&part.source)
        .map_err(|error| classify(error, part.source.len()))?;
    let changed_snapshot = dml::theme::family::part::read(&part.changed)
        .map_err(|error| classify(error, part.changed.len()))?;
    let changed_family = changed_snapshot
        .family()
        .cloned()
        .ok_or_else(|| fail("changed Theme part has no direct family owner"))?;
    let source_sha256 = sha256(&part.source);
    let replaced = theme_family_replace(&source_snapshot, &changed_family)
        .map_err(|error| classify(error, part.source.len()))?;
    if replaced == part.source {
        return Err(fail(
            "theme-family replace did not change the complete Theme part",
        ));
    }
    let removed =
        theme_family_remove(&part.source).map_err(|error| classify(error, part.source.len()))?;
    let removed_snapshot =
        dml::theme::family::part::read(&removed).map_err(|error| classify(error, removed.len()))?;
    if removed_snapshot.family().is_some() {
        return Err(fail("theme-family remove retained a direct owner"));
    }
    let added = theme_family_add(&removed, &changed_family)
        .map_err(|error| classify(error, removed.len()))?;
    let added_snapshot =
        dml::theme::family::part::read(&added).map_err(|error| classify(error, added.len()))?;
    if added_snapshot.family().is_none() {
        return Err(fail("theme-family add did not restore a direct owner"));
    }
    Ok(vec![
        ThemeFamilyMutationReceipt {
            operation: "replace_family",
            source_sha256: source_sha256.clone(),
            output_sha256: sha256(&replaced),
            output_bytes: replaced.len(),
            output_differs: true,
        },
        ThemeFamilyMutationReceipt {
            operation: "remove_family",
            source_sha256: source_sha256.clone(),
            output_sha256: sha256(&removed),
            output_bytes: removed.len(),
            output_differs: removed != part.source,
        },
        ThemeFamilyMutationReceipt {
            operation: "add_family",
            source_sha256: sha256(&removed),
            output_sha256: sha256(&added),
            output_bytes: added.len(),
            output_differs: added != removed,
        },
    ])
}

fn preflight_check(case: &LaneCase) -> Result<serde_json::Value, ProbeFailure> {
    let mut parts = Vec::with_capacity(case.parts.len());
    for part in &case.parts {
        let mut full_sink = CountingSink::new(usize::MAX);
        write_kind_to(part.kind, &part.changed, &mut full_sink)?;
        let full_receipt = write_receipt(full_sink);
        if full_receipt.bytes == 0 {
            return Err(fail(format!("{} preflight output was empty", part.id)));
        }
        let mut sink = CountingSink::new(full_receipt.bytes.saturating_sub(1));
        let error = write_kind_to(part.kind, &part.changed, &mut sink)
            .expect_err("bounded sink must refuse before its first write");
        let refusal = error
            .refusal
            .ok_or_else(|| fail(format!("{} preflight lacked a typed sink refusal", part.id)))?;
        if sink.writes != 0 {
            return Err(fail(format!("{} preflight wrote before refusal", part.id)));
        }
        parts.push(serde_json::json!({
            "part_id": part.id,
            "preflight_bytes": full_receipt.bytes,
            "sink_writes": sink.writes,
            "sink_limit": full_receipt.bytes - 1,
            "observed_output": full_receipt.bytes,
            "refusal": refusal,
        }));
    }
    Ok(serde_json::json!({
        "preflight_zero_writes": true,
        "sink_writes": 0,
        "parts": parts,
    }))
}

fn normal_write(case: &LaneCase) -> Result<SinkReceipt, ProbeFailure> {
    let mut sink = CountingSink::new(usize::MAX);
    for part in &case.parts {
        write_kind_to(part.kind, &part.changed, &mut sink)?;
    }
    Ok(write_receipt(sink))
}

fn vector_receipt(
    outputs: impl IntoIterator<Item = Result<Vec<u8>, ProbeFailure>>,
) -> Result<SinkReceipt, ProbeFailure> {
    let mut sink = CountingSink::new(usize::MAX);
    for output in outputs {
        let output = output?;
        sink.write_all(&output)
            .map_err(|error| io_failure(error, output.len()))?;
    }
    Ok(write_receipt(sink))
}

fn owner_operation(case: &LaneCase, operation_name: &str) -> Result<OperationValue, ProbeFailure> {
    let mut sink = CountingSink::new(usize::MAX);
    let mut refusals = Vec::new();
    for fixture in &case.owners {
        match operation_name {
            "noop_commit" => {
                let output = owner_commit(fixture, false)?.0;
                sink.write_all(&output)
                    .map_err(|error| io_failure(error, output.len()))?;
            },
            "changed_commit" => {
                let output = owner_commit(fixture, true)?.0;
                sink.write_all(&output)
                    .map_err(|error| io_failure(error, output.len()))?;
            },
            "forward_replay" => {
                let (changed, patch) = owner_commit(fixture, true)?;
                let _ = changed;
                let output = owner_replay(fixture, &patch, false)?;
                sink.write_all(&output)
                    .map_err(|error| io_failure(error, output.len()))?;
            },
            "inverse_replay" => {
                let (changed, patch) = owner_commit(fixture, true)?;
                let _ = changed;
                let output = owner_replay(fixture, &patch, true)?;
                sink.write_all(&output)
                    .map_err(|error| io_failure(error, output.len()))?;
            },
            "stale_replay" => {
                let (_changed, patch) = owner_commit(fixture, true)?;
                refusals.push(owner_stale_refusal(fixture, &patch)?);
            },
            _ => {
                return Err(fail(format!(
                    "unsupported owner operation {operation_name}"
                )));
            },
        }
    }
    if operation_name == "stale_replay" {
        Ok(OperationValue::Details(serde_json::json!({
            "refusal": refusals.first().cloned(),
            "refusals": refusals,
        })))
    } else {
        Ok(OperationValue::Sink(write_receipt(sink)))
    }
}

#[derive(Debug)]
enum OperationValue {
    Bytes(Vec<u8>),
    Sink(SinkReceipt),
    Details(serde_json::Value),
}

fn operation(case: &LaneCase, name: &str) -> Result<OperationValue, ProbeFailure> {
    if OWNER_OPERATIONS.contains(&name) {
        if case.owners.is_empty() {
            return Err(fail(format!(
                "owner operation {name} was requested for a raw-only lane"
            )));
        }
        return owner_operation(case, name);
    }
    match name {
        "entry_read" => Ok(OperationValue::Sink(vector_receipt(
            case.parts
                .iter()
                .map(|part| roundtrip(part.kind, &part.source)),
        )?)),
        "read_shared_when_supported" => {
            let outputs = read_shared_lane(case)?;
            if outputs.is_empty() {
                Ok(OperationValue::Details(serde_json::json!({
                    "not_applicable": true,
                    "reason": "no shared-reader API exists for this lane",
                })))
            } else {
                Ok(OperationValue::Sink(vector_receipt(
                    outputs.into_iter().map(Ok),
                )?))
            }
        },
        "vec_write" => Ok(OperationValue::Sink(vector_receipt(
            case.parts
                .iter()
                .map(|part| roundtrip(part.kind, &part.changed)),
        )?)),
        "write_to_bounded_sink" => Ok(OperationValue::Sink(normal_write(case)?)),
        "preflight_refusal" => Ok(OperationValue::Details(preflight_check(case)?)),
        "malformed_refusal" => Ok(OperationValue::Details(malformed_check(case)?)),
        _ => Err(fail(format!("unsupported operation {name}"))),
    }
}

fn aggregate_metrics(values: &[AllocationMetrics]) -> AllocationMetrics {
    let mut output = AllocationMetrics {
        conservation_ok: true,
        ..AllocationMetrics::default()
    };
    for value in values {
        output.allocation_calls = output
            .allocation_calls
            .saturating_add(value.allocation_calls);
        output.deallocation_calls = output
            .deallocation_calls
            .saturating_add(value.deallocation_calls);
        output.reallocation_calls = output
            .reallocation_calls
            .saturating_add(value.reallocation_calls);
        output.failed_calls = output.failed_calls.saturating_add(value.failed_calls);
        output.direct_bytes = output.direct_bytes.saturating_add(value.direct_bytes);
        output.deallocated_bytes = output
            .deallocated_bytes
            .saturating_add(value.deallocated_bytes);
        output.realloc_old_bytes = output
            .realloc_old_bytes
            .saturating_add(value.realloc_old_bytes);
        output.realloc_new_bytes = output
            .realloc_new_bytes
            .saturating_add(value.realloc_new_bytes);
        output.live_before = output.live_before.saturating_add(value.live_before);
        output.live_after = output.live_after.saturating_add(value.live_after);
        output.peak_live_delta = output.peak_live_delta.saturating_add(value.peak_live_delta);
        output.conservation_ok &= value.conservation_ok;
    }
    output.elapsed_claimable = false;
    output
}

fn process_rss_bytes() -> u64 {
    let Ok(status) = fs::read_to_string("/proc/self/status") else {
        return 0;
    };
    status
        .lines()
        .find_map(|line| {
            line.strip_prefix("VmRSS:")
                .and_then(|value| value.split_whitespace().next())
                .and_then(|value| value.parse::<u64>().ok())
                .map(|value| value.saturating_mul(1024))
        })
        .unwrap_or(0)
}

/// Return the kernel-maintained whole-process RSS high-water mark.  Unlike
/// the sample-side `VmRSS` observations, `VmHWM` is retained even when a
/// short allocation peak falls between samples.
fn whole_process_peak_rss_bytes() -> u64 {
    let Ok(status) = fs::read_to_string("/proc/self/status") else {
        return 0;
    };
    status
        .lines()
        .find_map(|line| {
            line.strip_prefix("VmHWM:")
                .and_then(|value| value.split_whitespace().next())
                .and_then(|value| value.parse::<u64>().ok())
                .map(|value| value.saturating_mul(1024))
        })
        .unwrap_or(0)
}

fn measure(
    case: &LaneCase,
    operation_name: &str,
    arguments: &Arguments,
    observer: Option<&dyn AllocationObserver>,
) -> Result<(Row, Option<AllocationMetrics>, u64), ProbeFailure> {
    let mut warm = Vec::new();
    for _ in 0..arguments.warmups {
        for _ in 0..arguments.iterations {
            warm.push(operation(case, operation_name)?);
        }
    }
    black_box(warm);
    let before = observer.map(AllocationObserver::begin);
    let mut samples = Vec::with_capacity(arguments.blocks * arguments.samples);
    let mut peak_rss = process_rss_bytes();
    let sample_target = Duration::from_millis(
        (arguments.target_ms / (arguments.blocks * arguments.samples) as u64).max(1),
    );
    let mut last_value = None;
    for _block in 0..arguments.blocks {
        for _sample in 0..arguments.samples {
            let start = Instant::now();
            let mut calls = 0usize;
            loop {
                for _ in 0..arguments.iterations {
                    last_value = Some(operation(case, operation_name)?);
                    calls = calls.saturating_add(1);
                }
                if start.elapsed() >= sample_target {
                    break;
                }
            }
            let elapsed_ns = start.elapsed().as_nanos();
            samples.push(SampleReceipt {
                elapsed_ns,
                iterations: calls,
            });
            peak_rss = peak_rss.max(process_rss_bytes());
        }
    }
    black_box(last_value);
    let elapsed_ns = samples.iter().map(|sample| sample.elapsed_ns).sum::<u128>();
    if elapsed_ns < u128::from(arguments.minimum_ms) * 1_000_000 {
        return Err(fail(format!(
            "{operation_name} leg was shorter than minimum duration"
        )));
    }
    let metrics = before.map(|before| observer.expect("observer disappeared").finish(before));
    let value = operation(case, operation_name)?;
    let (details, sink) = match value {
        OperationValue::Sink(sink) => (
            serde_json::json!({
                "sink_hash": sink.sha256,
                "sink_bytes": sink.bytes,
                "sink_writes": sink.writes
            }),
            Some(sink),
        ),
        OperationValue::Bytes(bytes) => (
            serde_json::json!({ "output_sha256": sha256(&bytes), "output_bytes": bytes.len() }),
            None,
        ),
        OperationValue::Details(details) => (details, None),
    };
    let typed_refusal_checked = matches!(
        operation_name,
        "stale_replay" | "preflight_refusal" | "malformed_refusal"
    );
    let fidelity_checked = true;
    Ok((
        Row {
            operation: operation_name.into(),
            status: "qualified",
            elapsed_ns,
            fidelity_checked,
            typed_refusal_checked,
            sink_writes: sink.as_ref().map_or(0, |value| value.writes),
            sink,
            samples,
            details,
        },
        metrics,
        peak_rss,
    ))
}

fn allocator_report(
    metrics: AllocationMetrics,
    process_rss_bytes: u64,
    peak_rss_bytes: u64,
    whole_process_peak_rss_bytes: u64,
) -> AllocatorReport {
    AllocatorReport {
        elapsed_claimable: metrics.elapsed_claimable,
        allocation_calls: metrics.allocation_calls,
        deallocation_calls: metrics.deallocation_calls,
        reallocation_calls: metrics.reallocation_calls,
        failed_calls: metrics.failed_calls,
        direct_bytes: metrics.direct_bytes,
        deallocated_bytes: metrics.deallocated_bytes,
        realloc_old_bytes: metrics.realloc_old_bytes,
        realloc_new_bytes: metrics.realloc_new_bytes,
        live_before: metrics.live_before,
        live_after: metrics.live_after,
        live_balance: i128::from(metrics.live_after) - i128::from(metrics.live_before),
        peak_live_delta: metrics.peak_live_delta,
        process_rss_bytes,
        peak_rss_bytes,
        whole_process_peak_rss_bytes,
        conservation_ok: metrics.conservation_ok,
    }
}

fn write_report(path: &Path, report: &Report) -> io::Result<()> {
    let parent = path
        .parent()
        .ok_or_else(|| invalid_argument("JSON output has no parent directory"))?;
    if !parent.is_dir() || parent.is_symlink() {
        return Err(invalid_argument(format!(
            "JSON output parent is not a regular directory: {parent:?}"
        )));
    }
    let file_name = path
        .file_name()
        .ok_or_else(|| invalid_argument("JSON output has no file name"))?;
    let temporary = parent.join(format!(".{}.tmp", file_name.to_string_lossy()));
    if temporary.exists() || temporary.is_symlink() {
        return Err(invalid_argument(format!(
            "temporary JSON output already exists: {temporary:?}"
        )));
    }
    let bytes =
        serde_json::to_vec_pretty(report).map_err(|error| io::Error::other(error.to_string()))?;
    let mut file = fs::File::create(&temporary)?;
    file.write_all(&bytes)?;
    file.write_all(b"\n")?;
    file.sync_all()?;
    fs::rename(temporary, path)
}

/// Run the real DrawingML adapter endpoint.
pub fn run(observer: Option<&dyn AllocationObserver>) -> io::Result<()> {
    let expected_mode = if observer.is_some() {
        "allocator"
    } else {
        "normal"
    };
    let arguments = parse(env::args().skip(1), expected_mode)
        .map_err(|error| io::Error::new(error.kind(), format!("{error}; {}", usage())))?;
    let (fixture_bytes, fixture) = load_fixture(&arguments.fixture, &arguments.lane)?;
    let mut case =
        make_case(&arguments.lane, &fixture).map_err(|error| io::Error::other(error.message))?;
    if arguments.endpoint == "owner" {
        select_owner_route(
            &mut case,
            arguments
                .owner_route
                .as_deref()
                .expect("owner route validated by argument parser"),
        )
        .map_err(|error| io::Error::other(error.message))?;
    } else if let Some(scope) = arguments.raw_scope.as_deref() {
        select_raw_scope(&mut case, scope).map_err(|error| io::Error::other(error.message))?;
    }
    let size_classes = if arguments.endpoint == "raw" {
        size_class_receipts(&case).map_err(|error| io::Error::other(error.message))?
    } else {
        Vec::new()
    };
    let raw_fidelity = if arguments.endpoint == "raw" {
        raw_fidelity(&case).map_err(|error| io::Error::other(error.message))?
    } else {
        Vec::new()
    };
    let theme_family_mutations =
        if arguments.endpoint == "raw" && arguments.raw_scope.as_deref() == Some("family") {
            theme_family_mutations(&case).map_err(|error| io::Error::other(error.message))?
        } else {
            Vec::new()
        };
    let owner_proofs = if arguments.endpoint == "owner" {
        case.owners
            .iter()
            .map(|fixture| owner_proof(&case.fixture_identity_sha256, fixture))
            .collect::<Result<Vec<_>, _>>()
            .map_err(|error| io::Error::other(error.message))?
    } else {
        Vec::new()
    };
    let operations = if arguments.endpoint == "raw" {
        RAW_OPERATIONS.to_vec()
    } else {
        OWNER_OPERATIONS.to_vec()
    };
    let mut rows = Vec::with_capacity(operations.len());
    let mut all_metrics = Vec::new();
    let mut peak_rss = process_rss_bytes();
    for operation_name in operations {
        let (row, metrics, row_peak) = measure(&case, &operation_name, &arguments, observer)
            .map_err(|error| io::Error::other(error.message))?;
        rows.push(row);
        if let Some(metrics) = metrics {
            all_metrics.push(metrics);
        }
        peak_rss = peak_rss.max(row_peak);
    }
    let metrics = aggregate_metrics(&all_metrics);
    let whole_process_peak_rss = whole_process_peak_rss_bytes();
    let report = Report {
        schema: "litchi.drawingml.raw.probe.v1",
        lane: arguments.lane,
        endpoint: arguments.endpoint,
        owner_route: arguments.owner_route,
        raw_scope: arguments.raw_scope,
        role: arguments.role,
        mode: arguments.mode,
        source_revision: arguments.source_revision,
        status: "qualified",
        qualification_claim: true,
        fixture_bytes: fixture_bytes.len(),
        fixture_sha256: sha256(&fixture_bytes),
        fixture_root_sha256: sha256(fixture.root.as_bytes()),
        fixture_mutation: fixture.mutation,
        fixture_lane_identity_sha256: case.fixture_identity_sha256.clone(),
        capabilities: capability_receipts(&case, &arguments.endpoint),
        raw_fidelity,
        theme_family_mutations,
        owner_proofs,
        size_classes,
        protocol: Protocol {
            warmups: arguments.warmups,
            samples_per_leg: arguments.samples,
            blocks: arguments.blocks,
            target_ms: arguments.target_ms,
            minimum_ms: arguments.minimum_ms,
            fresh_process: true,
            abba_order: ["baseline", "candidate", "candidate", "baseline"],
        },
        allocator_metrics: observer.map(|_| {
            allocator_report(
                metrics,
                process_rss_bytes(),
                peak_rss,
                whole_process_peak_rss,
            )
        }),
        resource_metrics: ResourceMetrics {
            process_rss_bytes: process_rss_bytes(),
            peak_rss_bytes: peak_rss,
            whole_process_peak_rss_bytes: whole_process_peak_rss,
        },
        rows,
        iterations: arguments.iterations,
    };
    write_report(&arguments.output, &report)
}
