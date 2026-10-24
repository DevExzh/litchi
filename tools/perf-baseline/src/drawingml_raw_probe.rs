//! Real source-backed DrawingML adapters for the paired raw qualification.
//!
//! This endpoint exercises the shared `litchi-drawingml` codecs directly and
//! leaves package placement to the owning format crates. Every qualified row
//! carries an outside-clock fidelity receipt. A missing adapter or typed
//! refusal is a hard error; it is never converted into a timing claim.

#![forbid(unsafe_code)]

use std::env;
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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
    production_sink: bool,
    production_boundary: &'static str,
}

fn supports_shared(kind: PartKind) -> bool {
    matches!(
        kind,
        PartKind::FormatCode | PartKind::ThemeFamily | PartKind::ThemeFamilyPart
    )
}

/// Only codecs which expose a real `Write`/`write_to` production endpoint are
/// admitted to the bounded-sink lane.  Vec/String codecs remain covered by
/// `entry_read` and `vec_write`; wrapping them in a synthetic sink would turn
/// an adapter detail into a false capability claim.
fn supports_production_sink(kind: PartKind) -> bool {
    matches!(
        kind,
        PartKind::Chart
            | PartKind::FormatCode
            | PartKind::SvgBlip
            | PartKind::Model3d
            | PartKind::ThemeFamily
            | PartKind::InkActions
            | PartKind::Transform
    )
}

fn boundary_kind(kind: PartKind) -> &'static str {
    match kind {
        PartKind::Geometry => "structural",
        _ if input_limit(kind).is_some() => "input_limit",
        _ => "none",
    }
}

#[cfg(any(
    feature = "drawingml-owner-current",
    feature = "drawingml-theme-family-current",
    feature = "drawingml-raw-current",
    feature = "drawingml-ink-limits-current"
))]
const fn structural_refusal_kind() -> &'static str {
    "limit"
}

#[cfg(not(any(
    feature = "drawingml-owner-current",
    feature = "drawingml-theme-family-current",
    feature = "drawingml-raw-current",
    feature = "drawingml-ink-limits-current"
)))]
const fn structural_refusal_kind() -> &'static str {
    "invalid_input"
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
                production_sink: supports_production_sink(part.kind),
                production_boundary: boundary_kind(part.kind),
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
                    production_sink: false,
                    production_boundary: "none",
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
    prepared: Vec<PreparedPart>,
    shared: Vec<SharedPart>,
    stale: Vec<u8>,
    owners: Vec<OwnerFixture>,
    prepared_owners: Vec<OwnerWork>,
    fixture_identity_sha256: String,
}

/// A parsed raw value retained outside the timed region.  The timed write
/// rows consume these values directly, so a writer measurement cannot hide a
/// second reader (or a decode/encode round trip) in its clock.
#[derive(Debug)]
enum PreparedValue {
    Chart(dml::chart::Chart),
    FormatCode(dml::chart::extension::formatcode2::Element),
    Diagram {
        source: Vec<u8>,
        outline: dml::diagram::edit::Outline,
    },
    TextBody(dml::text::body::Body),
    SvgBlip(dml::svg_blip::SvgBlip),
    Model3d(dml::model3d::Metadata),
    Blip(String),
    Theme(dml::theme::Theme),
    ThemeFamily(dml::theme::family::Family),
    ThemeFamilyPart(dml::theme::family::part::Snapshot),
    InkActions(dml::ink::actions::Profile),
    Geometry(dml::geometry::CustomGeometry),
    Color(dml::color::Value),
    Transform(dml::transform::Transform),
}

#[derive(Debug)]
struct PreparedPart {
    id: &'static str,
    source: PreparedValue,
    changed: PreparedValue,
}

#[derive(Debug)]
struct SharedPart {
    id: &'static str,
    kind: PartKind,
    source: std::sync::Arc<[u8]>,
}

#[derive(Debug, Clone)]
struct OwnerFixture {
    route: OwnerRoute,
    source: Vec<u8>,
    stale: Vec<u8>,
}

struct OwnerWork {
    fixture: OwnerFixture,
    changed: Vec<u8>,
    patch: OwnerPatch,
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
    source_preserving_expected: bool,
    source_output_exact: bool,
    fidelity_mode: &'static str,
    source_semantic_sha256: String,
    source_output_semantic_sha256: String,
    changed_semantic_sha256: String,
    changed_output_semantic_sha256: String,
    source_semantic_equal: bool,
    changed_semantic_equal: bool,
    semantic_changed_differs: bool,
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
    boundary: &'static str,
}

#[derive(Debug, Clone, Serialize)]
struct ResourceMetrics {
    pid: u32,
    command_sha256: String,
    source_revision: String,
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
    pid: u32,
    command_sha256: String,
    source_revision: String,
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

fn command_identity_sha256() -> String {
    let command = env::args().collect::<Vec<_>>().join("\0");
    sha256(command.as_bytes())
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

fn mce_limit_kind(kind: litchi_ooxml_common::mce::LimitKind) -> &'static str {
    use litchi_ooxml_common::mce::LimitKind;
    match kind {
        LimitKind::InputBytes => "mce_input_bytes",
        LimitKind::OutputBytes => "mce_output_bytes",
        LimitKind::SourceBytes => "mce_source_bytes",
        LimitKind::MarkedBytes => "mce_marked_bytes",
        LimitKind::OffsetCount => "mce_offset_count",
        LimitKind::MarkerBytes => "mce_marker_bytes",
        LimitKind::Depth => "mce_depth",
        LimitKind::NamespaceBindings => "mce_namespace_bindings",
        LimitKind::DirectiveTokens => "mce_directive_tokens",
        LimitKind::Choices => "mce_choices",
        LimitKind::Branches => "mce_branches",
        LimitKind::Attributes => "mce_attributes",
        LimitKind::Events => "mce_events",
        LimitKind::EventBytes => "mce_event_bytes",
        LimitKind::AttributeBytes => "mce_attribute_bytes",
        LimitKind::ContextBytes => "mce_context_bytes",
        LimitKind::NameBytes => "mce_name_bytes",
        LimitKind::ReportCounters => "mce_report_counters",
        LimitKind::RetryCount => "mce_retry_count",
        _ => "mce_unknown",
    }
}

#[cfg(not(any(
    feature = "drawingml-owner-current",
    feature = "drawingml-theme-family-current",
    feature = "drawingml-raw-current",
    feature = "drawingml-ink-limits-current"
)))]
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
        dml::Error::Mce(litchi_ooxml_common::mce::Error::LimitExceeded {
            kind,
            observed: error_observed,
            maximum,
        }) => {
            let resource = mce_limit_kind(kind);
            Some(RefusalReceipt {
                kind: "limit",
                source_kind: None,
                limit_kind: Some(resource),
                resource: Some(resource),
                limit: Some(maximum),
                observed: error_observed.max(observed),
                message: message.clone(),
            })
        },
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

#[cfg(any(
    feature = "drawingml-owner-current",
    feature = "drawingml-theme-family-current",
    feature = "drawingml-raw-current",
    feature = "drawingml-ink-limits-current"
))]
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
        dml::Error::Mce(litchi_ooxml_common::mce::Error::LimitExceeded {
            kind,
            observed: error_observed,
            maximum,
        }) => {
            let resource = mce_limit_kind(kind);
            Some(RefusalReceipt {
                kind: "limit",
                source_kind: None,
                limit_kind: Some(resource),
                resource: Some(resource),
                limit: Some(maximum),
                observed: error_observed.max(observed),
                message: message.clone(),
            })
        },
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

#[cfg(feature = "drawingml-raw-current")]
fn read_chart_value(bytes: &[u8]) -> Result<dml::chart::Chart, ProbeFailure> {
    dml::chart::reader::read(bytes, &dml::chart::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn read_chart_value(bytes: &[u8]) -> Result<dml::chart::Chart, ProbeFailure> {
    dml::chart::reader::read(bytes).map_err(|error| classify(error, bytes.len()))
}

#[cfg(feature = "drawingml-raw-current")]
fn write_chart_value<W: Write>(
    writer: &mut W,
    chart: &dml::chart::Chart,
) -> Result<(), ProbeFailure> {
    dml::chart::writer::write(writer, chart, &dml::chart::Limits::DEFAULT)
        .map_err(|error| classify(error, 0))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn write_chart_value<W: Write>(
    writer: &mut W,
    chart: &dml::chart::Chart,
) -> Result<(), ProbeFailure> {
    dml::chart::writer::write(writer, chart).map_err(|error| io_failure(error, 0))
}

fn read_chart(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let chart = read_chart_value(bytes)?;
    let mut output = Vec::new();
    write_chart_value(&mut output, &chart)?;
    Ok(output)
}

fn read_format_code_value(
    bytes: &[u8],
) -> Result<dml::chart::extension::formatcode2::Element, ProbeFailure> {
    #[cfg(feature = "drawingml-raw-current")]
    let result = dml::chart::extension::formatcode2::read(
        bytes,
        &dml::chart::extension::formatcode2::Limits::DEFAULT,
    );
    #[cfg(not(feature = "drawingml-raw-current"))]
    let result = dml::chart::extension::formatcode2::read(bytes);
    result.map_err(|error| classify(error, bytes.len()))
}

fn write_format_code_value<W: Write>(
    writer: &mut W,
    value: &dml::chart::extension::formatcode2::Element,
) -> Result<(), ProbeFailure> {
    #[cfg(feature = "drawingml-raw-current")]
    let result = dml::chart::extension::formatcode2::write_to(
        writer,
        value,
        &dml::chart::extension::formatcode2::Limits::DEFAULT,
    );
    #[cfg(not(feature = "drawingml-raw-current"))]
    let result = dml::chart::extension::formatcode2::write_to(writer, value);
    result.map_err(|error| classify(error, 0))
}

fn read_format_code(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let value = read_format_code_value(bytes)?;
    let mut output = Vec::new();
    write_format_code_value(&mut output, &value)?;
    Ok(output)
}

#[cfg(feature = "drawingml-raw-current")]
fn read_diagram_value(bytes: &[u8]) -> Result<dml::diagram::edit::Outline, ProbeFailure> {
    dml::diagram::edit::Outline::read(bytes, &dml::diagram::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn read_diagram_value(bytes: &[u8]) -> Result<dml::diagram::edit::Outline, ProbeFailure> {
    dml::diagram::edit::Outline::read(bytes).map_err(|error| classify(error, bytes.len()))
}

#[cfg(feature = "drawingml-raw-current")]
fn set_diagram_text(
    bytes: &[u8],
    id: &str,
    text: &str,
) -> Result<Vec<u8>, ProbeFailure> {
    dml::diagram::edit::set_text(bytes, id, text, &dml::diagram::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn set_diagram_text(bytes: &[u8], id: &str, text: &str) -> Result<Vec<u8>, ProbeFailure> {
    dml::diagram::edit::set_text(bytes, id, text)
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(feature = "drawingml-raw-current")]
fn set_diagram_reference(
    bytes: &[u8],
    relationship: Option<&str>,
) -> Result<Vec<u8>, ProbeFailure> {
    dml::diagram::edit::set_drawing_reference(
        bytes,
        relationship,
        &dml::diagram::Limits::DEFAULT,
    )
    .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn set_diagram_reference(
    bytes: &[u8],
    relationship: Option<&str>,
) -> Result<Vec<u8>, ProbeFailure> {
    dml::diagram::edit::set_drawing_reference(bytes, relationship)
        .map_err(|error| classify(error, bytes.len()))
}

fn read_diagram(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let outline = read_diagram_value(bytes)?;
    let entry = outline
        .entries
        .first()
        .ok_or_else(|| invalid_input("diagram data model has no content entry", bytes.len()))?;
    let output = set_diagram_text(bytes, &entry.id, &entry.text)?;
    set_diagram_reference(&output, outline.drawing_relationship.as_deref())
}

#[cfg(feature = "drawingml-raw-current")]
fn read_text_value(bytes: &[u8]) -> Result<dml::text::body::Body, ProbeFailure> {
    dml::text::body::reader::read(
        bytes,
        quick_xml::name::NamespaceResolver::default(),
        dml::text::body::reader::Limits::DEFAULT,
    )
    .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn read_text_value(bytes: &[u8]) -> Result<dml::text::body::Body, ProbeFailure> {
    dml::text::body::reader::read_with_options(
        bytes,
        quick_xml::name::NamespaceResolver::default(),
        dml::text::body::reader::Limits::DEFAULT,
    )
    .map_err(|error| classify(error, bytes.len()))
}

fn write_text_value<W: Write>(
    writer: &mut W,
    body: &dml::text::body::Body,
) -> Result<(), ProbeFailure> {
    let mut contents = String::new();
    #[cfg(feature = "drawingml-raw-current")]
    dml::text::body::writer::write(
        &mut contents,
        body,
        dml::text::body::writer::Limits::DEFAULT,
    )
    .map_err(|error| classify(error, 0))?;
    #[cfg(not(feature = "drawingml-raw-current"))]
    dml::text::body::writer::write(&mut contents, body)
        .map_err(|error| classify(error, 0))?;
    let mut output = String::with_capacity(contents.len() + 32);
    output.push_str("<a:txBody xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">");
    output.push_str(&contents);
    output.push_str("</a:txBody>");
    writer
        .write_all(output.as_bytes())
        .map_err(|error| io_failure(error, output.len()))
}

fn body_roundtrip(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let body = read_text_value(bytes)?;
    let mut output = Vec::new();
    write_text_value(&mut output, &body)?;
    Ok(output)
}

#[cfg(feature = "drawingml-raw-current")]
fn read_svg_value(bytes: &[u8]) -> Result<dml::svg_blip::SvgBlip, ProbeFailure> {
    dml::svg_blip::read(bytes, dml::svg_blip::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn read_svg_value(bytes: &[u8]) -> Result<dml::svg_blip::SvgBlip, ProbeFailure> {
    dml::svg_blip::read(bytes).map_err(|error| classify(error, bytes.len()))
}

fn write_svg_value<W: Write>(
    writer: &mut W,
    value: &dml::svg_blip::SvgBlip,
) -> Result<(), ProbeFailure> {
    #[cfg(feature = "drawingml-raw-current")]
    let result = dml::svg_blip::write_to(writer, value, dml::svg_blip::Limits::DEFAULT);
    #[cfg(not(feature = "drawingml-raw-current"))]
    let result = dml::svg_blip::write_to(writer, value);
    result.map_err(|error| classify(error, 0))
}

fn read_svg(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let value = read_svg_value(bytes)?;
    let mut output = Vec::new();
    write_svg_value(&mut output, &value)?;
    Ok(output)
}

#[cfg(feature = "drawingml-raw-current")]
fn read_svg_contextual_value(bytes: &[u8]) -> Result<dml::svg_blip::SvgBlip, ProbeFailure> {
    dml::svg_blip::read_contextual(
        bytes,
        &dml::svg_blip::NamespaceContext::empty(),
        dml::svg_blip::Limits::DEFAULT,
    )
    .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn read_svg_contextual_value(bytes: &[u8]) -> Result<dml::svg_blip::SvgBlip, ProbeFailure> {
    dml::svg_blip::read_contextual(bytes, &dml::svg_blip::NamespaceContext::empty())
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(feature = "drawingml-raw-current")]
fn write_svg_contextual_limited<W: Write>(
    writer: &mut W,
    value: &dml::svg_blip::SvgBlip,
    maximum: usize,
) -> Result<(), ProbeFailure> {
    let limits = dml::svg_blip::Limits::new(
        dml::svg_blip::Limits::DEFAULT.max_input_bytes,
        maximum,
    );
    dml::svg_blip::write_contextual_to(writer, value, limits)
        .map_err(|error| classify(error, maximum.saturating_add(1)))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn write_svg_contextual_limited<W: Write>(
    writer: &mut W,
    value: &dml::svg_blip::SvgBlip,
    maximum: usize,
) -> Result<(), ProbeFailure> {
    dml::svg_blip::write_contextual_to(writer, value, maximum)
        .map_err(|error| classify(error, maximum.saturating_add(1)))
}

#[cfg(feature = "drawingml-raw-current")]
fn write_svg_contextual_default(
    value: &dml::svg_blip::SvgBlip,
) -> Result<Vec<u8>, ProbeFailure> {
    dml::svg_blip::write_contextual(value, dml::svg_blip::Limits::DEFAULT)
        .map_err(|error| classify(error, 0))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn write_svg_contextual_default(
    value: &dml::svg_blip::SvgBlip,
) -> Result<Vec<u8>, ProbeFailure> {
    dml::svg_blip::write_contextual(value, dml::svg_blip::MAX_XML_BYTES)
        .map_err(|error| classify(error, 0))
}

#[cfg(feature = "drawingml-raw-current")]
fn read_model3d_value(bytes: &[u8]) -> Result<dml::model3d::Metadata, ProbeFailure> {
    dml::model3d::codec::read(bytes, dml::model3d::codec::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn read_model3d_value(bytes: &[u8]) -> Result<dml::model3d::Metadata, ProbeFailure> {
    dml::model3d::codec::read(bytes).map_err(|error| classify(error, bytes.len()))
}

fn write_model3d_value<W: Write>(
    writer: &mut W,
    value: &dml::model3d::Metadata,
) -> Result<(), ProbeFailure> {
    #[cfg(feature = "drawingml-raw-current")]
    let result = dml::model3d::codec::write_to(
        writer,
        value,
        dml::model3d::codec::Limits::DEFAULT,
    );
    #[cfg(not(feature = "drawingml-raw-current"))]
    let result = dml::model3d::codec::write_to(writer, value);
    result.map_err(|error| classify(error, 0))
}

fn read_model3d(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let value = read_model3d_value(bytes)?;
    let mut output = Vec::new();
    write_model3d_value(&mut output, &value)?;
    Ok(output)
}

#[cfg(feature = "drawingml-raw-current")]
fn read_blip_value(bytes: &[u8]) -> Result<String, ProbeFailure> {
    dml::blip::find_first_embed(bytes, dml::blip::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))?
        .ok_or_else(|| invalid_input("image blip has no embedded relationship", bytes.len()))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn read_blip_value(bytes: &[u8]) -> Result<String, ProbeFailure> {
    dml::blip::find_first_embed(bytes)
        .map_err(|error| classify(error, bytes.len()))?
        .ok_or_else(|| invalid_input("image blip has no embedded relationship", bytes.len()))
}

fn write_blip_value(value: &str) -> Result<Vec<u8>, ProbeFailure> {
    let mut output = String::new();
    dml::blip::write_embed(&mut output, value, true)
        .map_err(|error| invalid_input(error.to_string(), value.len()))?;
    Ok(output.into_bytes())
}

fn read_blip(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    write_blip_value(&read_blip_value(bytes)?)
}

#[cfg(feature = "drawingml-raw-current")]
fn read_theme_value(bytes: &[u8]) -> Result<dml::theme::Theme, ProbeFailure> {
    dml::theme::codec::read(bytes, dml::theme::codec::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn read_theme_value(bytes: &[u8]) -> Result<dml::theme::Theme, ProbeFailure> {
    dml::theme::codec::read(bytes).map_err(|error| classify(error, bytes.len()))
}

fn write_theme_value(value: &dml::theme::Theme) -> Result<Vec<u8>, ProbeFailure> {
    #[cfg(feature = "drawingml-raw-current")]
    let result = dml::theme::codec::encode_part(
        &value.name,
        &value.colors,
        &value.fonts,
        dml::theme::codec::Limits::DEFAULT,
    );
    #[cfg(not(feature = "drawingml-raw-current"))]
    let result = dml::theme::codec::encode_part(&value.name, &value.colors, &value.fonts);
    result.map_err(|error| classify(error, 0))
}

fn read_theme(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    write_theme_value(&read_theme_value(bytes)?)
}

#[cfg(any(
    feature = "drawingml-theme-family-current"
))]
fn read_family_value(bytes: &[u8]) -> Result<dml::theme::family::Family, ProbeFailure> {
    dml::theme::family::read(bytes, dml::theme::family::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(any(
    feature = "drawingml-theme-family-current"
)))]
fn read_family_value(bytes: &[u8]) -> Result<dml::theme::family::Family, ProbeFailure> {
    dml::theme::family::read(bytes).map_err(|error| classify(error, bytes.len()))
}

fn write_family_value(value: &dml::theme::family::Family) -> Result<Vec<u8>, ProbeFailure> {
    #[cfg(any(
        feature = "drawingml-theme-family-current"
    ))]
    let result = dml::theme::family::write(value, dml::theme::family::Limits::DEFAULT);
    #[cfg(not(any(
        feature = "drawingml-theme-family-current"
    )))]
    let result = dml::theme::family::write(value);
    result.map_err(|error| classify(error, 0))
}

fn read_family(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    write_family_value(&read_family_value(bytes)?)
}

#[cfg(any(
    feature = "drawingml-theme-family-current"
))]
fn read_theme_family_part_value(
    bytes: &[u8],
) -> Result<dml::theme::family::part::Snapshot, ProbeFailure> {
    dml::theme::family::part::read(bytes, dml::theme::family::part::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(any(
    feature = "drawingml-theme-family-current"
)))]
fn read_theme_family_part_value(
    bytes: &[u8],
) -> Result<dml::theme::family::part::Snapshot, ProbeFailure> {
    dml::theme::family::part::read(bytes).map_err(|error| classify(error, bytes.len()))
}

fn read_theme_family_part(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let snapshot = read_theme_family_part_value(bytes)?;
    let family = snapshot.family().ok_or_else(|| {
        invalid_input("theme family part has no direct family owner", bytes.len())
    })?;
    theme_family_replace(&snapshot, family).map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(any(
    feature = "drawingml-theme-family-current"
)))]
fn theme_family_replace(
    snapshot: &dml::theme::family::part::Snapshot,
    family: &dml::theme::family::Family,
) -> dml::Result<Vec<u8>> {
    snapshot.replace_family_with_limit(family, dml::theme::family::part::MAX_XML_BYTES)
}

#[cfg(any(
    feature = "drawingml-theme-family-current"
))]
fn theme_family_replace(
    snapshot: &dml::theme::family::part::Snapshot,
    family: &dml::theme::family::Family,
) -> dml::Result<Vec<u8>> {
    snapshot.replace_family(family, dml::theme::family::part::Limits::DEFAULT)
}

#[cfg(not(any(
    feature = "drawingml-theme-family-current"
)))]
fn theme_family_add(source: &[u8], family: &dml::theme::family::Family) -> dml::Result<Vec<u8>> {
    dml::theme::family::part::add_family_with_limit(
        source,
        family,
        dml::theme::family::part::MAX_XML_BYTES,
    )
}

#[cfg(any(
    feature = "drawingml-theme-family-current"
))]
fn theme_family_add(source: &[u8], family: &dml::theme::family::Family) -> dml::Result<Vec<u8>> {
    dml::theme::family::part::add_family(source, family, dml::theme::family::part::Limits::DEFAULT)
}

#[cfg(not(feature = "drawingml-theme-family-current"))]
fn theme_family_remove(source: &[u8]) -> dml::Result<Vec<u8>> {
    dml::theme::family::part::remove_family_with_limit(
        source,
        dml::theme::family::part::MAX_XML_BYTES,
    )
}

#[cfg(feature = "drawingml-theme-family-current")]
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

#[cfg(any(
    feature = "drawingml-raw-current",
    feature = "drawingml-ink-limits-current"
))]
fn read_ink_value(bytes: &[u8]) -> Result<dml::ink::actions::Profile, ProbeFailure> {
    dml::ink::actions::read_profile(bytes, dml::ink::actions::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(any(
    feature = "drawingml-raw-current",
    feature = "drawingml-ink-limits-current"
)))]
fn read_ink_value(bytes: &[u8]) -> Result<dml::ink::actions::Profile, ProbeFailure> {
    dml::ink::actions::read_profile(bytes).map_err(|error| classify(error, bytes.len()))
}

fn write_ink_value<W: Write>(
    writer: &mut W,
    value: &dml::ink::actions::Profile,
) -> Result<(), ProbeFailure> {
    #[cfg(any(
        feature = "drawingml-raw-current",
        feature = "drawingml-ink-limits-current"
    ))]
    let result = dml::ink::actions::write_profile_to(
        writer,
        value,
        dml::ink::actions::Limits::DEFAULT,
    );
    #[cfg(not(any(
        feature = "drawingml-raw-current",
        feature = "drawingml-ink-limits-current"
    )))]
    let result = dml::ink::actions::write_profile_to(writer, value);
    result.map_err(|error| classify(error, 0))
}

fn read_ink(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let profile = read_ink_value(bytes)?;
    let mut output = Vec::new();
    write_ink_value(&mut output, &profile)?;
    Ok(output)
}

fn read_geometry_value(
    bytes: &[u8],
) -> Result<dml::geometry::CustomGeometry, ProbeFailure> {
    parse_geometry(bytes)
}

fn write_geometry_value(
    geometry: &dml::geometry::CustomGeometry,
) -> Result<Vec<u8>, ProbeFailure> {
    let mut output = String::new();
    #[cfg(feature = "drawingml-raw-current")]
    dml::geometry::writer::write_custom_geometry(
        &mut output,
        geometry,
        &dml::geometry::Limits::DEFAULT,
    )
    .map_err(|error| classify(error, 0))?;
    #[cfg(not(feature = "drawingml-raw-current"))]
    dml::geometry::writer::write_custom_geometry(&mut output, geometry);
    Ok(output.into_bytes())
}

fn read_geometry(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    write_geometry_value(&read_geometry_value(bytes)?)
}

#[cfg(feature = "drawingml-raw-current")]
fn read_color_value(bytes: &[u8]) -> Result<dml::color::Value, ProbeFailure> {
    dml::color::read(bytes, &dml::color::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn read_color_value(bytes: &[u8]) -> Result<dml::color::Value, ProbeFailure> {
    dml::color::read(bytes).map_err(|error| classify(error, bytes.len()))
}

fn write_color_value(value: &dml::color::Value) -> Result<Vec<u8>, ProbeFailure> {
    #[cfg(feature = "drawingml-raw-current")]
    let result = dml::color::write(value, &dml::color::Limits::DEFAULT);
    #[cfg(not(feature = "drawingml-raw-current"))]
    let result = dml::color::write(value);
    result.map_err(|error| classify(error, 0))
}

fn read_color(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    write_color_value(&read_color_value(bytes)?)
}

#[cfg(feature = "drawingml-raw-current")]
fn read_transform_value(bytes: &[u8]) -> Result<dml::transform::Transform, ProbeFailure> {
    dml::transform::codec::read(bytes, &dml::transform::codec::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn read_transform_value(bytes: &[u8]) -> Result<dml::transform::Transform, ProbeFailure> {
    dml::transform::codec::read(bytes).map_err(|error| classify(error, bytes.len()))
}

fn write_transform_value<W: Write>(
    writer: &mut W,
    value: &dml::transform::Transform,
) -> Result<(), ProbeFailure> {
    #[cfg(feature = "drawingml-raw-current")]
    let result = dml::transform::codec::write_to(
        writer,
        value,
        &dml::transform::codec::Limits::DEFAULT,
    );
    #[cfg(not(feature = "drawingml-raw-current"))]
    let result = dml::transform::codec::write_to(writer, value);
    result.map_err(|error| classify(error, 0))
}

fn read_transform(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let value = read_transform_value(bytes)?;
    let mut output = Vec::new();
    write_transform_value(&mut output, &value)?;
    Ok(output)
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

fn prepare_value(kind: PartKind, bytes: &[u8]) -> Result<PreparedValue, ProbeFailure> {
    Ok(match kind {
        PartKind::Chart => PreparedValue::Chart(read_chart_value(bytes)?),
        PartKind::FormatCode => PreparedValue::FormatCode(read_format_code_value(bytes)?),
        PartKind::Diagram => PreparedValue::Diagram {
            source: bytes.to_vec(),
            outline: read_diagram_value(bytes)?,
        },
        PartKind::TextBody => PreparedValue::TextBody(read_text_value(bytes)?),
        PartKind::SvgBlip => PreparedValue::SvgBlip(read_svg_value(bytes)?),
        PartKind::Model3d => PreparedValue::Model3d(read_model3d_value(bytes)?),
        PartKind::Blip => PreparedValue::Blip(read_blip_value(bytes)?),
        PartKind::Theme => PreparedValue::Theme(read_theme_value(bytes)?),
        PartKind::ThemeFamily => PreparedValue::ThemeFamily(read_family_value(bytes)?),
        PartKind::ThemeFamilyPart => {
            PreparedValue::ThemeFamilyPart(read_theme_family_part_value(bytes)?)
        },
        PartKind::InkActions => PreparedValue::InkActions(read_ink_value(bytes)?),
        PartKind::Geometry => PreparedValue::Geometry(read_geometry_value(bytes)?),
        PartKind::Color => PreparedValue::Color(read_color_value(bytes)?),
        PartKind::Transform => PreparedValue::Transform(read_transform_value(bytes)?),
    })
}

fn write_prepared_to<W: Write>(
    kind: PartKind,
    value: &PreparedValue,
    writer: &mut W,
) -> Result<(), ProbeFailure> {
    match (kind, value) {
        (PartKind::Chart, PreparedValue::Chart(value)) => write_chart_value(writer, value),
        (PartKind::FormatCode, PreparedValue::FormatCode(value)) => {
            write_format_code_value(writer, value)
        },
        (PartKind::Diagram, PreparedValue::Diagram { source, outline }) => {
            let entry = outline
                .entries
                .first()
                .ok_or_else(|| fail("diagram data model has no content entry"))?;
            let output = set_diagram_text(source, &entry.id, &entry.text)?;
            let output = set_diagram_reference(&output, outline.drawing_relationship.as_deref())?;
            writer
                .write_all(&output)
                .map_err(|error| io_failure(error, output.len()))
        },
        (PartKind::TextBody, PreparedValue::TextBody(value)) => write_text_value(writer, value),
        (PartKind::SvgBlip, PreparedValue::SvgBlip(value)) => write_svg_value(writer, value),
        (PartKind::Model3d, PreparedValue::Model3d(value)) => write_model3d_value(writer, value),
        (PartKind::Blip, PreparedValue::Blip(value)) => {
            let output = write_blip_value(value)?;
            writer
                .write_all(&output)
                .map_err(|error| io_failure(error, output.len()))
        },
        (PartKind::Theme, PreparedValue::Theme(value)) => {
            let output = write_theme_value(value)?;
            writer
                .write_all(&output)
                .map_err(|error| io_failure(error, output.len()))
        },
        (PartKind::ThemeFamily, PreparedValue::ThemeFamily(value)) => {
            #[cfg(any(
                feature = "drawingml-raw-current",
                feature = "drawingml-theme-family-current"
            ))]
            let result = dml::theme::family::write_to(
                writer,
                value,
                dml::theme::family::Limits::DEFAULT,
            );
            #[cfg(not(any(
                feature = "drawingml-raw-current",
                feature = "drawingml-theme-family-current"
            )))]
            let result = dml::theme::family::write_to(writer, value);
            result.map_err(|error| classify(error, 0))
        },
        (PartKind::ThemeFamilyPart, PreparedValue::ThemeFamilyPart(snapshot)) => {
            let family = snapshot
                .family()
                .ok_or_else(|| fail("theme family part has no direct family owner"))?;
            let output = theme_family_replace(snapshot, family)
                .map_err(|error| classify(error, snapshot.xml_bytes().len()))?;
            writer
                .write_all(&output)
                .map_err(|error| io_failure(error, output.len()))
        },
        (PartKind::InkActions, PreparedValue::InkActions(value)) => {
            write_ink_value(writer, value)
        },
        (PartKind::Geometry, PreparedValue::Geometry(value)) => {
            let output = write_geometry_value(value)?;
            writer
                .write_all(&output)
                .map_err(|error| io_failure(error, output.len()))
        },
        (PartKind::Color, PreparedValue::Color(value)) => {
            let output = write_color_value(value)?;
            writer
                .write_all(&output)
                .map_err(|error| io_failure(error, output.len()))
        },
        (PartKind::Transform, PreparedValue::Transform(value)) => {
            write_transform_value(writer, value)
        },
        _ => Err(fail("prepared raw value kind mismatch")),
    }
}

fn write_prepared_vec(
    kind: PartKind,
    value: &PreparedValue,
) -> Result<Vec<u8>, ProbeFailure> {
    let mut output = Vec::new();
    write_prepared_to(kind, value, &mut output)?;
    Ok(output)
}

struct CountingSink {
    cap: usize,
    bytes: usize,
    writes: usize,
    digest: Option<Sha256>,
}

impl CountingSink {
    fn new(cap: usize) -> Self {
        Self::with_digest(cap, true)
    }

    fn timed(cap: usize) -> Self {
        Self::with_digest(cap, false)
    }

    fn with_digest(cap: usize, digest: bool) -> Self {
        Self {
            cap,
            bytes: 0,
            writes: 0,
            digest: digest.then(Sha256::new),
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
        if let Some(digest) = self.digest.as_mut() {
            digest.update(bytes);
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn write_receipt(sink: CountingSink) -> SinkReceipt {
    SinkReceipt {
        writes: sink.writes,
        bytes: sink.bytes,
        sha256: format!(
            "{:x}",
            sink.digest
                .expect("timed sink cannot produce a receipt")
                .finalize()
        ),
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
    write_chart_value(&mut source_wire, &chart)?;
    let source = read_chart(&source_wire)?;
    let changed_spec = Spec::new(Variant::new(Kind::Column))
        .with_title(format!("{} changed", fixture.mutation))
        .with_categories(["Q1", "Q2", "Q3"])
        .with_series(Series::new("East", vec![1.0, 2.0, 4.0]));
    let changed_chart = changed_spec
        .to_chart()
        .map_err(|error| classify(error, 0))?;
    let mut changed_wire = Vec::new();
    write_chart_value(&mut changed_wire, &changed_chart)?;
    let changed = read_chart(&changed_wire)?;
    Ok(vec![Part {
        id: "chart",
        kind: PartKind::Chart,
        source,
        changed,
        malformed: malformed_fragment("chart", &fixture.malformed),
        // chart::reader::read delegates its input admission to the shared
        // MCE default policy (256 MiB).  Keep the boundary fixture tied to
        // that real production limit rather than treating Chart as unbounded.
        over: over_limit(PartKind::Chart),
        limit_resource: "MCE input bytes",
    }])
}

fn format_code_part(fixture: &FixtureLane) -> Result<Part, ProbeFailure> {
    let mut value = dml::chart::extension::formatcode2::Element::new("0.00")
        .map_err(|error| classify(error, 0))?;
    let source = {
        let mut output = Vec::new();
        write_format_code_value(&mut output, &value)?;
        output
    };
    value
        .set_value(format!("0.00;{}", fixture.mutation.len()))
        .map_err(|error| classify(error, 0))?;
    let changed = {
        let mut output = Vec::new();
        write_format_code_value(&mut output, &value)?;
        output
    };
    Ok(Part {
        id: "formatcode2",
        kind: PartKind::FormatCode,
        source,
        changed,
        malformed: malformed_fragment("formatcode2", &fixture.malformed),
        over: over_limit(PartKind::FormatCode),
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
    let outline = read_diagram_value(&source)?;
    let entry = outline
        .entries
        .first()
        .ok_or_else(|| fail("diagram authoring produced no content entry"))?;
    let text_changed = set_diagram_text(
        &source,
        &entry.id,
        &format!("{} changed", fixture.mutation),
    )?;
    let changed = set_diagram_reference(&text_changed, Some("rIdDrawing"))?;
    Ok(vec![Part {
        id: "diagram-data-model",
        kind: PartKind::Diagram,
        source,
        changed,
        malformed: malformed_fragment("diagram-data-model", &fixture.malformed),
        over: over_limit(PartKind::Diagram),
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
    let mut output = Vec::new();
    write_text_value(&mut output, &body)?;
    Ok(output)
}

fn text_media_parts(fixture: &FixtureLane) -> Result<Vec<Part>, ProbeFailure> {
    let source = text_source()?;
    let mut body = read_text_value(&source)?;
    body.paragraphs[0].runs[0].text = format!("{} changed", fixture.mutation);
    let mut changed = Vec::new();
    write_text_value(&mut changed, &body)?;
    let svg_wire = br#"<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdSvg"/>"#.to_vec();
    let svg_source = read_svg(&svg_wire)?;
    let mut svg = read_svg_value(&svg_source)?;
    svg.set_reference(
        dml::svg_blip::Reference::embedded("rIdSvgChanged")
            .map_err(|error| fail(error.to_string()))?,
    )
    .map_err(|error| classify(error, svg_source.len()))?;
    let mut svg_changed = Vec::new();
    write_svg_value(&mut svg_changed, &svg)?;
    let model_wire = br#"<m3d:model3d xmlns:m3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdModel" r:link="rIdLinked"><m3d:spPr/><m3d:camera/><m3d:trans/><m3d:raster rName="Office Renderer" rVer="17.0"><m3d:blip r:embed="rIdRaster"/></m3d:raster><m3d:objViewport/></m3d:model3d>"#.to_vec();
    let model_source = read_model3d(&model_wire)?;
    let mut metadata = read_model3d_value(&model_source)?;
    metadata.reference = dml::model3d::Reference::embedded("rIdModelChanged")
        .map_err(|error| fail(error.to_string()))?;
    let mut model_changed = Vec::new();
    write_model3d_value(&mut model_changed, &metadata)?;
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
            over: over_limit(PartKind::TextBody),
            limit_resource: "text body input bytes",
        },
        Part {
            id: "svg-blip",
            kind: PartKind::SvgBlip,
            source: svg_source,
            changed: svg_changed,
            malformed: malformed_fragment("svg-blip", &fixture.malformed),
            over: over_limit(PartKind::SvgBlip),
            limit_resource: "SVG blip XML bytes",
        },
        Part {
            id: "model3d",
            kind: PartKind::Model3d,
            source: model_source,
            changed: model_changed,
            malformed: malformed_fragment("model3d", &fixture.malformed),
            over: over_limit(PartKind::Model3d),
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
    let source = write_theme_value(&dml::theme::Theme {
        name: "Office".to_owned(),
        colors: palette.clone(),
        fonts: fonts.clone(),
    })?;
    let changed = write_theme_value(&dml::theme::Theme {
        name: "Office Changed".to_owned(),
        colors: palette,
        fonts,
    })?;
    let family = dml::theme::family::Family::new(
        "Office",
        "{11111111-1111-1111-1111-111111111111}",
        "{22222222-2222-2222-2222-222222222222}",
    )
    .map_err(|error| classify(error, 0))?;
    let family_source = write_family_value(&family)?;
    let family_changed = family
        .with_name(format!("Family {}", fixture.mutation.len()))
        .map_err(|error| classify(error, family_source.len()))?;
    let family_changed = write_family_value(&family_changed)?;
    let family_part_source = complete_theme_family_part(&family_source)?;
    let family_part_changed = complete_theme_family_part(&family_changed)?;
    let ink_source = br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" lengthUnit="cm" timeUnit="s"><iact:action type="add" startTime="0.25"><iact:actionData name="stroke"/></iact:action><iact:actionGroup type="transform" startTime="1"><iact:action type="remove" startTime="2"/></iact:actionGroup></iact:actions>"#.to_vec();
    let ink_profile = read_ink_value(&ink_source)?;
    let ink_changed = ink_source
        .windows(4)
        .position(|window| window == b"0.25")
        .map(|position| {
            let mut output = ink_source.clone();
            output.splice(position..position + 4, b"0.50".iter().copied());
            output
        })
        .unwrap_or_else(|| ink_source.clone());
    let mut ink_output = Vec::new();
    write_ink_value(&mut ink_output, &ink_profile)?;
    Ok(vec![
        Part {
            id: "theme",
            kind: PartKind::Theme,
            source,
            changed,
            malformed: malformed_fragment("theme", &fixture.malformed),
            over: over_limit(PartKind::Theme),
            limit_resource: "theme XML bytes",
        },
        Part {
            id: "theme-family",
            kind: PartKind::ThemeFamily,
            source: family_source,
            changed: family_changed,
            malformed: malformed_fragment("theme-family", &fixture.malformed),
            over: over_limit(PartKind::ThemeFamily),
            limit_resource: "theme family XML bytes",
        },
        Part {
            id: "theme-family-part",
            kind: PartKind::ThemeFamilyPart,
            source: family_part_source,
            changed: family_part_changed,
            malformed: malformed_fragment("theme-family-part", &fixture.malformed),
            over: over_limit(PartKind::ThemeFamilyPart),
            limit_resource: "Theme Family XML bytes",
        },
        Part {
            id: "ink-actions",
            kind: PartKind::InkActions,
            source: ink_source,
            changed: ink_changed,
            malformed: malformed_fragment("ink-actions", &fixture.malformed),
            over: over_limit(PartKind::InkActions),
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
    let source = write_geometry_value(&geometry)?;
    let mut changed_geometry = geometry.clone();
    if let Some(PathCommand::LineTo(point)) = changed_geometry.paths[0].commands.get_mut(1) {
        point.x = AdjustValue::Value(90);
    }
    let changed_geometry_xml = write_geometry_value(&changed_geometry)?;
    // The reader admits at most 1024 paths.  This is a structural boundary,
    // not an input-byte policy: the bytes remain small enough to exercise the
    // actual geometry parser before it allocates a 1025th path.
    let mut structural_over = String::from(
        r#"<a:custGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:pathLst>"#,
    );
    for _ in 0..1025 {
        structural_over.push_str(r#"<a:path w="1" h="1"/>"#);
    }
    structural_over.push_str("</a:pathLst></a:custGeom>");
    let color_source = write_color_value(&dml::color::Value::rgb(
        dml::color::Rgb::new(51, 102, 153),
    ))?;
    let color_changed = write_color_value(&dml::color::Value::rgb(
        dml::color::Rgb::new(153, 102, 51),
    ))?;
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
    let transform_source = {
        let mut output = Vec::new();
        write_transform_value(&mut output, &transform)?;
        output
    };
    let transform_changed = {
        let mut output = Vec::new();
        write_transform_value(
            &mut output,
            &transform.with_rotation(dml::transform::Angle::new(60000)),
        )?;
        output
    };
    Ok(vec![
        Part {
            id: "geometry",
            kind: PartKind::Geometry,
            source,
            changed: changed_geometry_xml,
            malformed: malformed_fragment("geometry", &fixture.malformed),
            over: structural_over.into_bytes(),
            limit_resource: "geometry structural limits",
        },
        Part {
            id: "color",
            kind: PartKind::Color,
            source: color_source,
            changed: color_changed,
            malformed: malformed_fragment("color", &fixture.malformed),
            over: over_limit(PartKind::Color),
            limit_resource: "DrawingML color XML",
        },
        format_part,
        Part {
            id: "transform",
            kind: PartKind::Transform,
            source: transform_source,
            changed: transform_changed,
            malformed: malformed_fragment("transform", &fixture.malformed),
            over: over_limit(PartKind::Transform),
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
    let prepared = parts
        .iter()
        .map(|part| {
            Ok(PreparedPart {
                id: part.id,
                source: prepare_value(part.kind, &part.source)?,
                changed: prepare_value(part.kind, &part.changed)?,
            })
        })
        .collect::<Result<Vec<_>, ProbeFailure>>()?;
    let shared = parts
        .iter()
        .filter(|part| supports_shared(part.kind))
        .map(|part| SharedPart {
            id: part.id,
            kind: part.kind,
            source: std::sync::Arc::<[u8]>::from(part.source.clone()),
        })
        .collect();
    Ok(LaneCase {
        parts,
        prepared,
        shared,
        stale,
        owners,
        prepared_owners: Vec::new(),
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
    let fixture = &case.owners[0];
    let (changed, patch) = owner_commit(fixture, true)?;
    case.prepared_owners.push(OwnerWork {
        fixture: fixture.clone(),
        changed,
        patch,
    });
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
    let selected = case
        .parts
        .iter()
        .map(|part| part.id)
        .collect::<std::collections::HashSet<_>>();
    case.prepared.retain(|part| selected.contains(part.id));
    case.shared.retain(|part| selected.contains(part.id));
    case.owners.clear();
    case.prepared_owners.clear();
    Ok(())
}

#[cfg(not(feature = "drawingml-owner-current"))]
fn owner_error(route: OwnerRoute, error: dml::Error) -> ProbeFailure {
    let message = error.to_string();
    let refusal = if legacy_patch_conflict(route, &error) {
        Some(RefusalReceipt {
            kind: "patch_conflict",
            source_kind: Some("invalid_input"),
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
fn legacy_patch_conflict(route: OwnerRoute, error: &dml::Error) -> bool {
    let dml::Error::Invalid(message) = error else {
        return false;
    };
    match route {
        OwnerRoute::DrawingmlChart => {
            message == "chart patch source does not match its byte precondition"
        },
        OwnerRoute::DrawingmlThemeFamily => {
            message == "theme family patch source does not match its byte precondition"
        },
        OwnerRoute::DrawingmlInkActions => message == "ink action patch source is stale",
        OwnerRoute::DrawingmlTransform => {
            message == "DrawingML transform patch source state does not match its precondition"
        },
    }
}

#[cfg(feature = "drawingml-owner-current")]
fn owner_error(_route: OwnerRoute, error: dml::Error) -> ProbeFailure {
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
            use dml::ink::actions::{ActionSelector, Edit, EditOptions};
            let profile = dml::ink::actions::read_profile(&fixture.source)
                .map_err(|error| owner_error(fixture.route, error))?;
            let mut edit = Edit::from_profile(&profile, EditOptions::default())
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
            use dml::ink::actions::{Edit, EditOptions};
            let patch = if inverse {
                patch.inverse()
            } else {
                patch.clone()
            };
            let profile = dml::ink::actions::read_profile(source)
                .map_err(|error| owner_error(route, error))?;
            let prepared = Edit::from_profile(&profile, EditOptions::default())
                .map_err(|error| owner_error(route, error))?
                .commit()
                .map_err(|error| owner_error(route, error))?
                .into_package();
            prepared
                .apply(&patch)
                .map(|commit| commit.package().as_bytes().to_vec())
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
                let Some(mut refusal) = error.refusal else {
                    return Err(fail(format!(
                        "{} stale replay returned an unexpected error: {}",
                        owner_route_name(fixture.route),
                        error.message
                    )));
                };
                refusal.observed = fixture.stale.len();
                Ok(refusal)
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

fn semantic_mode(kind: PartKind) -> &'static str {
    match kind {
        PartKind::FormatCode
        | PartKind::SvgBlip
        | PartKind::ThemeFamily
        | PartKind::ThemeFamilyPart
        | PartKind::InkActions => "source_exact",
        PartKind::Blip => "scanner",
        _ => "model",
    }
}

fn semantic_key(kind: PartKind, bytes: &[u8]) -> Result<String, ProbeFailure> {
    Ok(match kind {
        PartKind::Chart => format!("{:?}", read_chart_value(bytes)?),
        PartKind::FormatCode => read_format_code_value(bytes)?.value().to_owned(),
        PartKind::Diagram => format!("{:?}", read_diagram_value(bytes)?),
        PartKind::TextBody => {
            let body = read_text_value(bytes)?;
            format!("text={:?};paragraphs={}", body.text(), body.paragraphs.len())
        },
        PartKind::SvgBlip => {
            let value = read_svg_value(bytes)?;
            format!(
                "reference={:?};prefix={:?};namespaces={:?};attributes={:?};children={:?}",
                value.reference(),
                value.prefix(),
                value.namespaces(),
                value.attributes(),
                value.children(),
            )
        },
        PartKind::Model3d => format!("{:?}", read_model3d_value(bytes)?),
        PartKind::Blip => read_blip_value(bytes)?,
        PartKind::Theme => format!("{:?}", read_theme_value(bytes)?),
        PartKind::ThemeFamily => {
            let value = read_family_value(bytes)?;
            format!("{}|{}|{}", value.name(), value.id(), value.variant_id())
        },
        PartKind::ThemeFamilyPart => {
            let value = read_theme_family_part_value(bytes)?;
            let family = value
                .family()
                .map(|family| format!("{}|{}|{}", family.name(), family.id(), family.variant_id()));
            format!("has_family={};uri={:?};family={family:?}", value.has_family(), value.family_extension_uri())
        },
        PartKind::InkActions => {
            let value = read_ink_value(bytes)?;
            format!(
                "id={:?};length={};time={};children={:?}",
                value.xml_id(),
                value.length_unit().as_str(),
                value.time_unit().as_str(),
                value.children(),
            )
        },
        PartKind::Geometry => format!("{:?}", read_geometry_value(bytes)?),
        PartKind::Color => format!("{:?}", read_color_value(bytes)?),
        PartKind::Transform => format!("{:?}", read_transform_value(bytes)?),
    })
}

fn semantic_equal(kind: PartKind, left: &[u8], right: &[u8]) -> Result<bool, ProbeFailure> {
    Ok(match kind {
        PartKind::Chart => format!("{:?}", read_chart_value(left)?)
            == format!("{:?}", read_chart_value(right)?),
        PartKind::FormatCode => read_format_code_value(left)? == read_format_code_value(right)?,
        PartKind::Diagram => read_diagram_value(left)? == read_diagram_value(right)?,
        PartKind::TextBody => read_text_value(left)? == read_text_value(right)?,
        PartKind::SvgBlip => semantic_key(kind, left)? == semantic_key(kind, right)?,
        PartKind::Model3d => read_model3d_value(left)? == read_model3d_value(right)?,
        PartKind::Blip => read_blip_value(left)? == read_blip_value(right)?,
        PartKind::Theme => read_theme_value(left)? == read_theme_value(right)?,
        PartKind::ThemeFamily => semantic_key(kind, left)? == semantic_key(kind, right)?,
        PartKind::ThemeFamilyPart => semantic_key(kind, left)? == semantic_key(kind, right)?,
        PartKind::InkActions => semantic_key(kind, left)? == semantic_key(kind, right)?,
        PartKind::Geometry => read_geometry_value(left)? == read_geometry_value(right)?,
        PartKind::Color => read_color_value(left)? == read_color_value(right)?,
        PartKind::Transform => read_transform_value(left)? == read_transform_value(right)?,
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
            let source_semantic = semantic_key(part.kind, &part.source)?;
            let source_output_semantic = semantic_key(part.kind, &source_output)?;
            let changed_semantic = semantic_key(part.kind, &part.changed)?;
            let changed_output_semantic = semantic_key(part.kind, &changed_output)?;
            let source_semantic_equal =
                semantic_equal(part.kind, &part.source, &source_output)?;
            let changed_semantic_equal =
                semantic_equal(part.kind, &part.changed, &changed_output)?;
            let semantic_changed_differs =
                !semantic_equal(part.kind, &part.source, &part.changed)?;
            if !source_semantic_equal || !changed_semantic_equal {
                return Err(fail(format!(
                    "{} raw write changed the decoded semantic model",
                    part.id
                )));
            }
            if !semantic_changed_differs {
                return Err(fail(format!(
                    "{} changed fixture did not change its semantic model",
                    part.id
                )));
            }
            let source_preserving_expected = source_preserving(part.kind);
            if source_preserving_expected && source_output != part.source {
                return Err(fail(format!(
                    "{} canonical source changed under a promised source-preserving route",
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
                changed_output_differs: source_output != changed_output,
                source_preserving_expected,
                source_output_exact: source_output == part.source,
                fidelity_mode: semantic_mode(part.kind),
                source_semantic_sha256: sha256(source_semantic.as_bytes()),
                source_output_semantic_sha256: sha256(source_output_semantic.as_bytes()),
                changed_semantic_sha256: sha256(changed_semantic.as_bytes()),
                changed_output_semantic_sha256: sha256(changed_output_semantic.as_bytes()),
                source_semantic_equal,
                changed_semantic_equal,
                semantic_changed_differs,
            })
        })
        .collect()
}

fn source_preserving(kind: PartKind) -> bool {
    matches!(
        kind,
        PartKind::FormatCode
            | PartKind::SvgBlip
            | PartKind::ThemeFamily
            | PartKind::ThemeFamilyPart
            | PartKind::InkActions
    )
}

fn size_class_receipts(case: &LaneCase) -> Result<Vec<SizeClassReceipt>, ProbeFailure> {
    let mut receipts = Vec::with_capacity(case.parts.len().saturating_mul(5));
    for part in &case.parts {
        let boundary = boundary_kind(part.kind);
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
        } else if boundary == "structural" {
            vec![
                ("small", part.source.clone(), None, true),
                ("near_limit", part.source.clone(), None, true),
                ("at_limit", part.source.clone(), None, true),
                (
                    "over_limit",
                    part.over.clone(),
                    Some(structural_refusal_kind()),
                    true,
                ),
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
                    boundary,
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
                        boundary,
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
                        boundary,
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
        PartKind::Chart => Some(("MCE input bytes", 256 * 1024 * 1024)),
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
        PartKind::Blip | PartKind::Geometry => None,
    }
}

fn over_limit(kind: PartKind) -> Vec<u8> {
    let (_, limit) = input_limit(kind).expect("fixture boundary requires an input limit");
    vec![b' '; limit.saturating_add(1)]
}

fn part_has_input_limit(kind: PartKind) -> bool {
    input_limit(kind).is_some()
}

fn parse_geometry(bytes: &[u8]) -> Result<dml::geometry::CustomGeometry, ProbeFailure> {
    use dml::geometry::reader::{CustomGeometryBuilder, GeometryElement};
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    #[cfg(feature = "drawingml-raw-current")]
    let mut builder = CustomGeometryBuilder::new(dml::geometry::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))?;
    #[cfg(not(feature = "drawingml-raw-current"))]
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

fn read_shared_value(part: &SharedPart) -> Result<PreparedValue, ProbeFailure> {
    match part.kind {
        PartKind::FormatCode => {
            #[cfg(feature = "drawingml-raw-current")]
            let value = dml::chart::extension::formatcode2::read_shared(
                std::sync::Arc::clone(&part.source),
                &dml::chart::extension::formatcode2::Limits::DEFAULT,
            );
            #[cfg(not(feature = "drawingml-raw-current"))]
            let value = dml::chart::extension::formatcode2::read_shared(
                std::sync::Arc::clone(&part.source),
            );
            Ok(PreparedValue::FormatCode(
                value.map_err(|error| classify(error, part.source.len()))?,
            ))
        },
        PartKind::ThemeFamily => {
            #[cfg(feature = "drawingml-theme-family-current")]
            let value = dml::theme::family::read_shared(
                std::sync::Arc::clone(&part.source),
                dml::theme::family::Limits::DEFAULT,
            );
            #[cfg(not(feature = "drawingml-theme-family-current"))]
            let value = dml::theme::family::read_shared(std::sync::Arc::clone(&part.source));
            Ok(PreparedValue::ThemeFamily(
                value.map_err(|error| classify(error, part.source.len()))?,
            ))
        },
        PartKind::ThemeFamilyPart => {
            #[cfg(feature = "drawingml-theme-family-current")]
            let value = dml::theme::family::part::read_shared(
                std::sync::Arc::clone(&part.source),
                dml::theme::family::part::Limits::DEFAULT,
            );
            #[cfg(not(feature = "drawingml-theme-family-current"))]
            let value = dml::theme::family::part::read_shared(std::sync::Arc::clone(&part.source));
            Ok(PreparedValue::ThemeFamilyPart(
                value.map_err(|error| classify(error, part.source.len()))?,
            ))
        },
        _ => Err(fail(format!("{} has no shared-reader route", part.id))),
    }
}

fn read_shared_lane(case: &LaneCase) -> Result<Vec<PreparedValue>, ProbeFailure> {
    case.shared.iter().map(read_shared_value).collect()
}

fn shared_operation(case: &LaneCase, timed: bool) -> Result<OperationValue, ProbeFailure> {
    if timed {
        for value in read_shared_lane(case)? {
            black_box(value);
        }
        return Ok(OperationValue::Timed);
    }
    let mut sink = CountingSink::new(usize::MAX);
    for value in read_shared_lane(case)? {
        let output = match &value {
            PreparedValue::FormatCode(value) => {
                let mut output = Vec::new();
                write_format_code_value(&mut output, value)?;
                output
            },
            PreparedValue::ThemeFamily(value) => write_family_value(value)?,
            PreparedValue::ThemeFamilyPart(value) => value.xml_bytes().to_vec(),
            _ => return Err(fail("shared-reader returned an unsupported value")),
        };
        sink.write_all(&output)
            .map_err(|error| io_failure(error, output.len()))?;
    }
    Ok(OperationValue::Sink(write_receipt(sink)))
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
        let boundary = boundary_kind(part.kind);
        if boundary == "none" {
            continue;
        }
        let expected_kind = if boundary == "structural" {
            structural_refusal_kind()
        } else {
            "limit"
        };
        let error = roundtrip(part.kind, &part.over)
            .map(|_| ())
            .expect_err("boundary source must be refused");
        let refusal = error
            .refusal
            .ok_or_else(|| fail(format!("{} boundary input lacked typed refusal", part.id)))?;
        if refusal.kind != expected_kind {
            return Err(fail(format!(
                "{} boundary source was not a {} refusal: {}",
                part.id, expected_kind, error.message
            )));
        }
        if expected_kind == "limit"
            && (refusal.limit_kind.is_none()
                || refusal.limit.is_none()
                || refusal.observed <= refusal.limit.unwrap_or(0))
        {
            return Err(fail(format!(
                "{} boundary source was not a typed limit refusal: {}",
                part.id, error.message
            )));
        }
        parts.push(serde_json::json!({
            "part_id": part.id,
            "boundary": boundary,
            "expected_kind": expected_kind,
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
    let source_snapshot = read_theme_family_part_value(&part.source)?;
    let changed_snapshot = read_theme_family_part_value(&part.changed)?;
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
    let removed_snapshot = read_theme_family_part_value(&removed)?;
    if removed_snapshot.family().is_some() {
        return Err(fail("theme-family remove retained a direct owner"));
    }
    let added = theme_family_add(&removed, &changed_family)
        .map_err(|error| classify(error, removed.len()))?;
    let added_snapshot = read_theme_family_part_value(&added)?;
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
    let mut total_sink_writes = 0usize;
    for part in &case.parts {
        let boundary = boundary_kind(part.kind);
        let production_refusal = if boundary != "none" {
            let error = roundtrip(part.kind, &part.over)
                .map(|_| ())
                .expect_err("production boundary source must be refused");
            let refusal = error.refusal.ok_or_else(|| {
                fail(format!(
                    "{} production boundary lacked a typed refusal",
                    part.id
                ))
            })?;
            let expected = if boundary == "structural" {
                structural_refusal_kind()
            } else {
                "limit"
            };
            if refusal.kind != expected {
                return Err(fail(format!(
                    "{} production boundary returned {} instead of {expected}",
                    part.id, refusal.kind
                )));
            }
            Some(refusal)
        } else {
            None
        };
        if !supports_production_sink(part.kind) {
            parts.push(serde_json::json!({
                "part_id": part.id,
                "boundary": boundary,
                "production_refusal": production_refusal,
                "sink_supported": false,
                "sink_writes": 0,
            }));
            continue;
        }
        let prepared = case
            .prepared
            .iter()
            .find(|value| value.id == part.id)
            .ok_or_else(|| fail(format!("missing prepared value for {}", part.id)))?;
        if part.kind != PartKind::SvgBlip {
            parts.push(serde_json::json!({
                "part_id": part.id,
                "boundary": boundary,
                "production_refusal": production_refusal,
                "sink_supported": true,
                "preflight": "not_applicable",
                "reason": "the codec exposes a sink but no caller-selected output preflight",
                "sink_writes": 0,
            }));
            continue;
        }
        // Use the contextual production reader for this preflight receipt.  A
        // host embedding the fragment owns the inherited namespace scope;
        // the empty scope models a standalone root while still exercising the
        // same contextual writer and its complete-output admission check.
        let contextual = read_svg_contextual_value(&part.changed)?;
        let value = &contextual;
        let PreparedValue::SvgBlip(_) = &prepared.changed else {
            return Err(fail("prepared SVG value has the wrong kind"));
        };
        let full_output = write_svg_contextual_default(value)?;
        if full_output.is_empty() {
            return Err(fail(format!("{} preflight output was empty", part.id)));
        }
        // The cap is passed to the production SVG codec.  The sink is
        // deliberately unbounded: it observes writes but never supplies the
        // refusal boundary itself.
        let mut sink = CountingSink::new(usize::MAX);
        let error = write_svg_contextual_limited(
            &mut sink,
            value,
            full_output.len().saturating_sub(1),
        )
        .expect_err("SVG production preflight must refuse before its first write");
        let refusal = error
            .refusal
            .ok_or_else(|| fail(format!("{} preflight lacked a typed sink refusal", part.id)))?;
        if sink.writes != 0 {
            return Err(fail(format!("{} preflight wrote before refusal", part.id)));
        }
        total_sink_writes = total_sink_writes.saturating_add(sink.writes);
        parts.push(serde_json::json!({
            "part_id": part.id,
            "boundary": boundary,
            "production_refusal": production_refusal,
            "sink_supported": true,
            "preflight": "typed_output_limit",
            "preflight_bytes": full_output.len(),
            "sink_writes": sink.writes,
            "output_limit": full_output.len() - 1,
            "observed_output": full_output.len(),
            "sink_refusal": refusal,
        }));
    }
    Ok(serde_json::json!({
        "preflight_zero_writes": total_sink_writes == 0,
        "sink_writes": total_sink_writes,
        "parts": parts,
    }))
}

/// Execute only production boundary calls during a timed preflight row.
/// Receipt construction and output-cap experiments stay outside the clock.
fn production_preflight(case: &LaneCase) -> Result<(), ProbeFailure> {
    for part in &case.parts {
        if boundary_kind(part.kind) == "none" {
            continue;
        }
        let error = roundtrip(part.kind, &part.over)
            .map(|_| ())
            .expect_err("production boundary source must be refused");
        black_box(error);
    }
    for part in &case.parts {
        if part.kind != PartKind::SvgBlip {
            continue;
        }
        let prepared = case
            .prepared
            .iter()
            .find(|value| value.id == part.id)
            .ok_or_else(|| fail(format!("missing prepared value for {}", part.id)))?;
        let PreparedValue::SvgBlip(value) = &prepared.changed else {
            return Err(fail("prepared SVG value has the wrong kind"));
        };
        let full_output = write_svg_contextual_default(value)?;
        let mut sink = CountingSink::new(usize::MAX);
        let error = write_svg_contextual_limited(
            &mut sink,
            value,
            full_output.len().saturating_sub(1),
        )
        .expect_err("SVG production preflight must refuse before its first write");
        black_box(error);
    }
    Ok(())
}

/// The timed malformed row calls each real reader/writer route.  It does not
/// serialize a JSON receipt or calculate a digest until the outside-clock
/// detail pass.
fn malformed_production(case: &LaneCase) -> Result<(), ProbeFailure> {
    for part in &case.parts {
        let error = roundtrip(part.kind, &part.malformed)
            .map(|_| ())
            .expect_err("malformed source must be refused");
        black_box(error);
    }
    Ok(())
}

fn normal_write(case: &LaneCase, timed: bool) -> Result<OperationValue, ProbeFailure> {
    let production_parts = case
        .parts
        .iter()
        .filter(|part| supports_production_sink(part.kind))
        .collect::<Vec<_>>();
    if production_parts.is_empty() {
        return Ok(if timed {
            OperationValue::Timed
        } else {
            OperationValue::Details(serde_json::json!({
                "not_applicable": true,
                "reason": "this lane exposes only Vec/String raw writers",
                "production_sink_parts": [],
            }))
        });
    }
    let mut sink = if timed {
        CountingSink::timed(usize::MAX)
    } else {
        CountingSink::new(usize::MAX)
    };
    for part in production_parts {
        let prepared = case
            .prepared
            .iter()
            .find(|value| value.id == part.id)
            .ok_or_else(|| fail(format!("missing prepared value for {}", part.id)))?;
        write_prepared_to(part.kind, &prepared.changed, &mut sink)?;
    }
    if timed {
        Ok(OperationValue::Timed)
    } else {
        let sink = write_receipt(sink);
        Ok(OperationValue::SinkDetails {
            sink,
            details: serde_json::json!({
                "production_sink_parts": case
                    .parts
                    .iter()
                    .filter(|part| supports_production_sink(part.kind))
                    .map(|part| part.id)
                    .collect::<Vec<_>>(),
            }),
        })
    }
}

/// Read timing owns only the reader call.  The outside-clock receipt pass
/// performs the corresponding write so it can retain a digest without adding
/// serialization work to the read sample.
fn read_entry_value(kind: PartKind, bytes: &[u8]) -> Result<(), ProbeFailure> {
    match kind {
        PartKind::Chart => {
            black_box(read_chart_value(bytes)?);
        },
        PartKind::FormatCode => {
            black_box(read_format_code_value(bytes)?);
        },
        PartKind::Diagram => {
            black_box(read_diagram_value(bytes)?);
        },
        PartKind::TextBody => {
            black_box(read_text_value(bytes)?);
        },
        PartKind::SvgBlip => {
            black_box(read_svg_value(bytes)?);
        },
        PartKind::Model3d => {
            black_box(read_model3d_value(bytes)?);
        },
        PartKind::Blip => {
            black_box(read_blip_value(bytes)?);
        },
        PartKind::Theme => {
            black_box(read_theme_value(bytes)?);
        },
        PartKind::ThemeFamily => {
            black_box(read_family_value(bytes)?);
        },
        PartKind::ThemeFamilyPart => {
            black_box(read_theme_family_part_value(bytes)?);
        },
        PartKind::InkActions => {
            black_box(read_ink_value(bytes)?);
        },
        PartKind::Geometry => {
            black_box(read_geometry_value(bytes)?);
        },
        PartKind::Color => {
            black_box(read_color_value(bytes)?);
        },
        PartKind::Transform => {
            black_box(read_transform_value(bytes)?);
        },
    }
    Ok(())
}

fn entry_read_operation(case: &LaneCase, timed: bool) -> Result<OperationValue, ProbeFailure> {
    if timed {
        for part in &case.parts {
            read_entry_value(part.kind, &part.source)?;
        }
        return Ok(OperationValue::Timed);
    }
    let mut sink = CountingSink::new(usize::MAX);
    for part in &case.parts {
        let value = prepare_value(part.kind, &part.source)?;
        let output = write_prepared_vec(part.kind, &value)?;
        sink.write_all(&output)
            .map_err(|error| io_failure(error, output.len()))?;
    }
    Ok(OperationValue::Sink(write_receipt(sink)))
}

/// Write timing consumes values parsed once in `make_case`; no reader or
/// source round trip is hidden in the writer sample.
fn vec_write_operation(case: &LaneCase, timed: bool) -> Result<OperationValue, ProbeFailure> {
    if timed {
        for part in &case.parts {
            let prepared = case
                .prepared
                .iter()
                .find(|value| value.id == part.id)
                .ok_or_else(|| fail(format!("missing prepared value for {}", part.id)))?;
            black_box(write_prepared_vec(part.kind, &prepared.changed)?);
        }
        return Ok(OperationValue::Timed);
    }
    let mut sink = CountingSink::new(usize::MAX);
    for part in &case.parts {
        let prepared = case
            .prepared
            .iter()
            .find(|value| value.id == part.id)
            .ok_or_else(|| fail(format!("missing prepared value for {}", part.id)))?;
        let output = write_prepared_vec(part.kind, &prepared.changed)?;
        sink.write_all(&output)
            .map_err(|error| io_failure(error, output.len()))?;
    }
    Ok(OperationValue::Sink(write_receipt(sink)))
}

fn owner_operation(
    case: &LaneCase,
    operation_name: &str,
    timed: bool,
) -> Result<OperationValue, ProbeFailure> {
    let mut sink = if timed {
        CountingSink::timed(usize::MAX)
    } else {
        CountingSink::new(usize::MAX)
    };
    let mut refusals = Vec::new();
    for work in &case.prepared_owners {
        let fixture = &work.fixture;
        match operation_name {
            "noop_commit" => {
                let output = owner_commit(fixture, false)?.0;
                if timed {
                    black_box(output);
                } else {
                    sink.write_all(&output)
                        .map_err(|error| io_failure(error, output.len()))?;
                }
            },
            "changed_commit" => {
                let output = owner_commit(fixture, true)?.0;
                if timed {
                    black_box(output);
                } else {
                    sink.write_all(&output)
                        .map_err(|error| io_failure(error, output.len()))?;
                }
            },
            "forward_replay" => {
                let output = owner_apply(fixture.route, &fixture.source, &work.patch, false)?;
                if timed {
                    black_box(output);
                } else {
                    sink.write_all(&output)
                        .map_err(|error| io_failure(error, output.len()))?;
                }
            },
            "inverse_replay" => {
                let output = owner_apply(fixture.route, &work.changed, &work.patch, true)?;
                if timed {
                    black_box(output);
                } else {
                    sink.write_all(&output)
                        .map_err(|error| io_failure(error, output.len()))?;
                }
            },
            "stale_replay" => {
                if timed {
                    let error = owner_apply(fixture.route, &fixture.stale, &work.patch, false)
                        .expect_err("owner route accepted a stale patch");
                    black_box(error);
                } else {
                    refusals.push(owner_stale_refusal(fixture, &work.patch)?);
                }
            },
            _ => {
                return Err(fail(format!(
                    "unsupported owner operation {operation_name}"
                )));
            },
        }
    }
    if operation_name == "stale_replay" {
        Ok(if timed {
            OperationValue::Timed
        } else {
            OperationValue::Details(serde_json::json!({
                "refusal": refusals.first().cloned(),
                "refusals": refusals,
            }))
        })
    } else if timed {
        Ok(OperationValue::Timed)
    } else {
        Ok(OperationValue::Sink(write_receipt(sink)))
    }
}

#[derive(Debug)]
enum OperationValue {
    Timed,
    Sink(SinkReceipt),
    SinkDetails {
        sink: SinkReceipt,
        details: serde_json::Value,
    },
    Details(serde_json::Value),
}

fn operation(case: &LaneCase, name: &str, timed: bool) -> Result<OperationValue, ProbeFailure> {
    if OWNER_OPERATIONS.contains(&name) {
        if case.owners.is_empty() {
            return Err(fail(format!(
                "owner operation {name} was requested for a raw-only lane"
            )));
        }
        return owner_operation(case, name, timed);
    }
    match name {
        "entry_read" => entry_read_operation(case, timed),
        "read_shared_when_supported" => {
            if !case.parts.iter().any(|part| supports_shared(part.kind)) {
                Ok(if timed {
                    OperationValue::Timed
                } else {
                    OperationValue::Details(serde_json::json!({
                        "not_applicable": true,
                        "reason": "no shared-reader API exists for this lane",
                    }))
                })
            } else {
                shared_operation(case, timed)
            }
        },
        "vec_write" => vec_write_operation(case, timed),
        "write_to_bounded_sink" => normal_write(case, timed),
        "preflight_refusal" => {
            if timed {
                production_preflight(case)?;
                Ok(OperationValue::Timed)
            } else {
                Ok(OperationValue::Details(preflight_check(case)?))
            }
        },
        "malformed_refusal" => {
            if timed {
                malformed_production(case)?;
                Ok(OperationValue::Timed)
            } else {
                Ok(OperationValue::Details(malformed_check(case)?))
            }
        },
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
    outside_fidelity_checked: bool,
) -> Result<(Row, Option<AllocationMetrics>, u64), ProbeFailure> {
    for _ in 0..arguments.warmups {
        for _ in 0..arguments.iterations {
            black_box(operation(case, operation_name, true)?);
        }
    }
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
                    last_value = Some(operation(case, operation_name, true)?);
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
    let value = operation(case, operation_name, false)?;
    let (details, sink) = match value {
        OperationValue::Timed => (
            serde_json::json!({
                "not_applicable": true,
                "reason": "operation produced no outside-clock receipt",
            }),
            None,
        ),
        OperationValue::Sink(sink) => (
            serde_json::json!({
                "sink_hash": sink.sha256,
                "sink_bytes": sink.bytes,
                "sink_writes": sink.writes
            }),
            Some(sink),
        ),
        OperationValue::SinkDetails { sink, details } => {
            let details = serde_json::json!({
                "sink_hash": sink.sha256,
                "sink_bytes": sink.bytes,
                "sink_writes": sink.writes,
                "route": details,
            });
            (details, Some(sink))
        },
        OperationValue::Details(details) => (details, None),
    };
    let typed_refusal_checked = matches!(
        operation_name,
        "stale_replay" | "preflight_refusal" | "malformed_refusal"
    );
    // The detail pass (including semantic/source and replay checks) is always
    // outside the timed samples.  This marker is derived from the non-empty
    // receipt set produced before timing, rather than asserting a capability
    // without evidence.
    let fidelity_checked = outside_fidelity_checked;
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
    source_revision: &str,
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
        pid: std::process::id(),
        command_sha256: command_identity_sha256(),
        source_revision: source_revision.to_owned(),
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
        let (row, metrics, row_peak) = measure(
            &case,
            &operation_name,
            &arguments,
            observer,
            !raw_fidelity.is_empty() || !owner_proofs.is_empty(),
        )
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
        source_revision: arguments.source_revision.clone(),
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
                &arguments.source_revision,
            )
        }),
        resource_metrics: ResourceMetrics {
            pid: std::process::id(),
            command_sha256: command_identity_sha256(),
            source_revision: arguments.source_revision.clone(),
            process_rss_bytes: process_rss_bytes(),
            peak_rss_bytes: peak_rss,
            whole_process_peak_rss_bytes: whole_process_peak_rss,
        },
        rows,
        iterations: arguments.iterations,
    };
    write_report(&arguments.output, &report)
}
