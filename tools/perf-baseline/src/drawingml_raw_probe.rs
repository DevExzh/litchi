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
    preflight: bool,
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
            "crates/litchi-drawingml/src/ink/actions_edit.rs:Prepared::edit (current) or Edit::from_profile (legacy),Edit::commit,Prepared/Patch::apply"
        },
        OwnerRoute::DrawingmlTransform => {
            "crates/litchi-drawingml/src/transform/transaction.rs:Package::edit,Edit::commit,Package::apply"
        },
    }
}

#[cfg(not(feature = "drawingml-owner-current"))]
fn owner_operation_path(route: OwnerRoute, operation: &str) -> &'static str {
    match (route, operation) {
        (
            OwnerRoute::DrawingmlChart
            | OwnerRoute::DrawingmlThemeFamily
            | OwnerRoute::DrawingmlTransform,
            "noop_commit" | "changed_commit",
        ) => "Snapshot::edit -> Edit::commit",
        (OwnerRoute::DrawingmlInkActions, "noop_commit" | "changed_commit") => {
            "Edit::from_profile -> Edit::commit"
        },
        (_, "forward_replay" | "inverse_replay" | "stale_replay") => "Patch::apply",
        _ => "unknown owner operation",
    }
}

#[cfg(feature = "drawingml-owner-current")]
fn owner_operation_path(route: OwnerRoute, operation: &str) -> &'static str {
    match (route, operation) {
        (
            OwnerRoute::DrawingmlChart
            | OwnerRoute::DrawingmlThemeFamily
            | OwnerRoute::DrawingmlTransform,
            "noop_commit" | "changed_commit",
        ) => "Package::edit -> Edit::commit",
        (OwnerRoute::DrawingmlInkActions, "noop_commit" | "changed_commit") => {
            "Prepared::edit -> Edit::commit"
        },
        (
            OwnerRoute::DrawingmlChart
            | OwnerRoute::DrawingmlThemeFamily
            | OwnerRoute::DrawingmlTransform,
            "forward_replay" | "inverse_replay" | "stale_replay",
        ) => "Package::apply",
        (OwnerRoute::DrawingmlInkActions, "forward_replay" | "inverse_replay" | "stale_replay") => {
            "Prepared::apply"
        },
        _ => "unknown owner operation",
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
            "crates/litchi-drawingml/src/text/body/reader.rs:read(xml,scope,limits);writer"
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

struct LaneCase {
    parts: Vec<Part>,
    prepared: Vec<PreparedPart>,
    shared: Vec<SharedPart>,
    shared_prepared: Vec<PreparedValue>,
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
        /// The source-edit target is selected before the timed writer.  The
        /// timed route must not walk `entries` just to recover its first
        /// content node.
        entry_id: String,
        entry_text: String,
        drawing_relationship: Option<String>,
    },
    TextBody(dml::text::body::Body),
    SvgBlip(dml::svg_blip::SvgBlip),
    Model3d(dml::model3d::Metadata),
    Blip(String),
    Theme(dml::theme::Theme),
    ThemeFamily(dml::theme::family::Family),
    ThemeFamilyPart {
        snapshot: dml::theme::family::part::Snapshot,
        /// The recognized owner is selected before timing.  Complete Theme
        /// mutations are source-edit plus serialization, not a detached
        /// writer, but the family lookup itself is still preparation.
        family: dml::theme::family::Family,
    },
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
    /// The contextual SVG value is prepared once before timing.  It is kept
    /// separate from the standalone value used by the ordinary Vec writer so
    /// the production preflight and its detail receipt exercise the same
    /// contextual object.
    contextual_source: Option<dml::svg_blip::SvgBlip>,
    contextual_changed: Option<dml::svg_blip::SvgBlip>,
    contextual_namespace: Option<dml::svg_blip::NamespaceContext>,
    contextual_source_output_bytes: Option<usize>,
    contextual_changed_output_bytes: Option<usize>,
    output_payload_bytes: Option<usize>,
    output_limit: Option<usize>,
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
    inverse_patch: OwnerPatch,
    prepared: OwnerPrepared,
    changed_prepared: OwnerPrepared,
    stale_prepared: OwnerPrepared,
}

#[derive(Debug)]
enum OwnerPatch {
    DrawingmlChart(dml::chart::edit::Patch),
    DrawingmlThemeFamily(dml::theme::family::Patch),
    DrawingmlInkActions(dml::ink::actions::Patch),
    DrawingmlTransform(dml::transform::Patch),
}

#[cfg(not(feature = "drawingml-owner-current"))]
enum OwnerPrepared {
    DrawingmlChart(dml::chart::edit::Snapshot),
    DrawingmlThemeFamily(dml::theme::family::Snapshot),
    DrawingmlInkActions(dml::ink::actions::Profile),
    DrawingmlTransform(dml::transform::Snapshot),
}

#[cfg(feature = "drawingml-owner-current")]
enum OwnerPrepared {
    DrawingmlChart(dml::chart::edit::Package),
    DrawingmlThemeFamily(dml::theme::family::Package),
    DrawingmlInkActions(dml::ink::actions::Prepared),
    DrawingmlTransform(dml::transform::Package),
}

#[derive(Debug, Clone, Serialize)]
struct OwnerSemanticReceipt {
    source_semantic_sha256: String,
    noop_semantic_sha256: String,
    changed_semantic_sha256: String,
    forward_semantic_sha256: String,
    inverse_semantic_sha256: String,
    source_rest_semantic_sha256: String,
    changed_rest_semantic_sha256: String,
    expected_mutation: &'static str,
    mutation_verified: bool,
    unchanged_remainder: bool,
}

#[derive(Debug, Clone, Serialize)]
struct OwnerCommitReceipt {
    route: &'static str,
    native_commit_path: &'static str,
    native_replay_path: &'static str,
    fixture_identity_sha256: String,
    source_sha256: String,
    noop_sha256: String,
    stale_sha256: String,
    changed_sha256: String,
    forward_sha256: String,
    inverse_sha256: String,
    stale_kind: &'static str,
    stale_error_variant: &'static str,
    stale_reason_verified: bool,
    stale_refusal: RefusalReceipt,
    semantic: OwnerSemanticReceipt,
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
    changed_output_exact: bool,
    /// Contextual SVG keeps the inherited fragment as `raw_source`; its
    /// exported output necessarily inserts the inherited declarations.  The
    /// raw-source receipt therefore proves exact source retention while the
    /// changed exported bytes remain a dirty contextual serialization.
    source_raw_sha256: String,
    source_raw_exact: bool,
    changed_output_exact_expected: bool,
    /// The contextual route is the measured normal path for inherited host
    /// fragments.  This separate flag records the additional standalone
    /// reader/writer check on the completed contextual output.
    standalone_route_checked: bool,
    read_write_route: &'static str,
    fidelity_mode: &'static str,
    source_semantic_sha256: String,
    source_output_semantic_sha256: String,
    changed_semantic_sha256: String,
    changed_output_semantic_sha256: String,
    source_semantic_equal: bool,
    changed_semantic_equal: bool,
    semantic_changed_differs: bool,
    /// Independent source-slice/splice evidence for the non-canonical
    /// source-backed formatcode2 and Theme-family fixtures.  This is kept
    /// separate from production read/write equality so a generated output
    /// cannot serve as its own oracle.
    splice_oracle_checked: bool,
    splice_oracle_passed: bool,
}

#[derive(Debug, Clone, Serialize)]
struct GeneratedFixturePart {
    part_id: &'static str,
    source_sha256: String,
    source_bytes: usize,
    changed_sha256: String,
    changed_bytes: usize,
    malformed_sha256: String,
    malformed_bytes: usize,
}

#[derive(Debug)]
struct OutputLimitAttempt {
    refusal: RefusalReceipt,
    sink: Option<SinkReceipt>,
    partial_output_allowed: bool,
}

#[derive(Debug, Clone, Serialize)]
struct ThemeFamilyMutationReceipt {
    operation: &'static str,
    source_sha256: String,
    output_sha256: String,
    output_bytes: usize,
    output_differs: bool,
    requested_name: Option<String>,
    requested_id: Option<String>,
    requested_variant_id: Option<String>,
    observed_name: Option<String>,
    observed_id: Option<String>,
    observed_variant_id: Option<String>,
    requested_identity_matches: bool,
    source_name: Option<String>,
    source_id: Option<String>,
    source_variant_id: Option<String>,
    source_identity_matches: bool,
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
    generated_fixture_sha256: String,
    generated_fixture_parts: Vec<GeneratedFixturePart>,
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
     --blocks N --samples N --target-ms N --minimum-ms N --json PATH [--preflight]"
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
    let mut preflight = false;
    let mut arguments = arguments;
    while let Some(argument) = arguments.next() {
        if argument == "--preflight" {
            if preflight {
                return Err(invalid_argument("duplicate argument --preflight"));
            }
            preflight = true;
            continue;
        }
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
    if preflight && expected_mode != "normal" {
        return Err(invalid_argument(
            "--preflight is only supported by the normal binary",
        ));
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
        preflight,
    })
}

fn sha256(bytes: &[u8]) -> String {
    let mut digest = Sha256::new();
    digest.update(bytes);
    hex_digest(digest.finalize())
}

fn hex_digest<D: AsRef<[u8]>>(digest: D) -> String {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    let bytes = digest.as_ref();
    let mut output = String::with_capacity(bytes.len() * 2);
    for &byte in bytes {
        output.push(HEX[(byte >> 4) as usize] as char);
        output.push(HEX[(byte & 0x0f) as usize] as char);
    }
    output
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
fn set_diagram_text(bytes: &[u8], id: &str, text: &str) -> Result<Vec<u8>, ProbeFailure> {
    dml::diagram::edit::set_text(bytes, id, text, &dml::diagram::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn set_diagram_text(bytes: &[u8], id: &str, text: &str) -> Result<Vec<u8>, ProbeFailure> {
    dml::diagram::edit::set_text(bytes, id, text).map_err(|error| classify(error, bytes.len()))
}

#[cfg(feature = "drawingml-raw-current")]
fn set_diagram_reference(
    bytes: &[u8],
    relationship: Option<&str>,
) -> Result<Vec<u8>, ProbeFailure> {
    dml::diagram::edit::set_drawing_reference(bytes, relationship, &dml::diagram::Limits::DEFAULT)
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
    dml::text::body::writer::write(&mut contents, body).map_err(|error| classify(error, 0))?;
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
    read_svg_contextual_value(bytes)
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn read_svg_value(bytes: &[u8]) -> Result<dml::svg_blip::SvgBlip, ProbeFailure> {
    read_svg_contextual_value(bytes)
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

#[cfg(feature = "drawingml-raw-current")]
fn read_svg_standalone_value(bytes: &[u8]) -> Result<dml::svg_blip::SvgBlip, ProbeFailure> {
    dml::svg_blip::read(bytes, dml::svg_blip::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn read_svg_standalone_value(bytes: &[u8]) -> Result<dml::svg_blip::SvgBlip, ProbeFailure> {
    dml::svg_blip::read(bytes).map_err(|error| classify(error, bytes.len()))
}

/// Run the standalone SVG API only on a completed contextual document.  The
/// inherited-fragment fixture itself must use `read_contextual`; this extra
/// route proves that the normal self-contained reader/writer remains covered
/// without changing the timed contextual workload.
fn standalone_svg_roundtrip(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let value = read_svg_standalone_value(bytes)?;
    let mut output = Vec::new();
    write_svg_value(&mut output, &value)?;
    Ok(output)
}

fn read_svg(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    let value = read_svg_value(bytes)?;
    write_svg_contextual_default(&value)
}

/// The SVG lane models a fragment selected from a host XML part.  Its root
/// intentionally omits declarations and relies on this inherited scope; a
/// fresh empty context would only prove the self-contained fast path.
fn svg_fixture_context() -> Result<dml::svg_blip::NamespaceContext, ProbeFailure> {
    use dml::svg_blip::{Namespace, NamespaceContext};
    NamespaceContext::empty()
        .child([
            Namespace::new(Some("asvg"), dml::svg_blip::NAMESPACE)
                .map_err(|error| fail(error.to_string()))?,
            Namespace::new(Some("r"), dml::svg_blip::RELATIONSHIP_NAMESPACE)
                .map_err(|error| fail(error.to_string()))?,
            Namespace::new(Some("future"), "urn:litchi:drawingml:future")
                .map_err(|error| fail(error.to_string()))?,
            Namespace::new(Some("opaque"), "urn:litchi:drawingml:opaque")
                .map_err(|error| fail(error.to_string()))?,
        ])
        .map_err(|error| fail(error.to_string()))
}

#[cfg(feature = "drawingml-raw-current")]
fn read_svg_contextual_value_with(
    bytes: &[u8],
    context: &dml::svg_blip::NamespaceContext,
) -> Result<dml::svg_blip::SvgBlip, ProbeFailure> {
    dml::svg_blip::read_contextual(bytes, context, dml::svg_blip::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn read_svg_contextual_value_with(
    bytes: &[u8],
    context: &dml::svg_blip::NamespaceContext,
) -> Result<dml::svg_blip::SvgBlip, ProbeFailure> {
    dml::svg_blip::read_contextual(bytes, context).map_err(|error| classify(error, bytes.len()))
}

#[cfg(feature = "drawingml-raw-current")]
fn read_svg_contextual_value(bytes: &[u8]) -> Result<dml::svg_blip::SvgBlip, ProbeFailure> {
    let context = svg_fixture_context()?;
    read_svg_contextual_value_with(bytes, &context)
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn read_svg_contextual_value(bytes: &[u8]) -> Result<dml::svg_blip::SvgBlip, ProbeFailure> {
    let context = svg_fixture_context()?;
    read_svg_contextual_value_with(bytes, &context)
}

#[cfg(feature = "drawingml-raw-current")]
fn write_svg_contextual_limited<W: Write>(
    writer: &mut W,
    value: &dml::svg_blip::SvgBlip,
    maximum: usize,
) -> Result<(), ProbeFailure> {
    let limits =
        dml::svg_blip::Limits::new(dml::svg_blip::Limits::DEFAULT.max_input_bytes, maximum);
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
fn write_svg_contextual_default(value: &dml::svg_blip::SvgBlip) -> Result<Vec<u8>, ProbeFailure> {
    dml::svg_blip::write_contextual(value, dml::svg_blip::Limits::DEFAULT)
        .map_err(|error| classify(error, 0))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn write_svg_contextual_default(value: &dml::svg_blip::SvgBlip) -> Result<Vec<u8>, ProbeFailure> {
    dml::svg_blip::write_contextual(value, dml::svg_blip::MAX_XML_BYTES)
        .map_err(|error| classify(error, 0))
}

#[cfg(feature = "drawingml-raw-current")]
fn write_svg_contextual_to_default<W: Write>(
    writer: &mut W,
    value: &dml::svg_blip::SvgBlip,
) -> Result<(), ProbeFailure> {
    dml::svg_blip::write_contextual_to(writer, value, dml::svg_blip::Limits::DEFAULT)
        .map_err(|error| classify(error, 0))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn write_svg_contextual_to_default<W: Write>(
    writer: &mut W,
    value: &dml::svg_blip::SvgBlip,
) -> Result<(), ProbeFailure> {
    dml::svg_blip::write_contextual_to(writer, value, dml::svg_blip::MAX_XML_BYTES)
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
    let result = dml::model3d::codec::write_to(writer, value, dml::model3d::codec::Limits::DEFAULT);
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

#[cfg(feature = "drawingml-raw-current")]
fn read_family_value(bytes: &[u8]) -> Result<dml::theme::family::Family, ProbeFailure> {
    dml::theme::family::read(bytes, dml::theme::family::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn read_family_value(bytes: &[u8]) -> Result<dml::theme::family::Family, ProbeFailure> {
    dml::theme::family::read(bytes).map_err(|error| classify(error, bytes.len()))
}

fn write_family_value(value: &dml::theme::family::Family) -> Result<Vec<u8>, ProbeFailure> {
    #[cfg(feature = "drawingml-raw-current")]
    let result = dml::theme::family::write(value, dml::theme::family::Limits::DEFAULT);
    #[cfg(not(feature = "drawingml-raw-current"))]
    let result = dml::theme::family::write(value);
    result.map_err(|error| classify(error, 0))
}

fn read_family(bytes: &[u8]) -> Result<Vec<u8>, ProbeFailure> {
    write_family_value(&read_family_value(bytes)?)
}

#[cfg(feature = "drawingml-raw-current")]
fn read_theme_family_part_value(
    bytes: &[u8],
) -> Result<dml::theme::family::part::Snapshot, ProbeFailure> {
    dml::theme::family::part::read(bytes, dml::theme::family::part::Limits::DEFAULT)
        .map_err(|error| classify(error, bytes.len()))
}

#[cfg(not(feature = "drawingml-raw-current"))]
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

#[cfg(not(feature = "drawingml-theme-family-current"))]
fn theme_family_replace(
    snapshot: &dml::theme::family::part::Snapshot,
    family: &dml::theme::family::Family,
) -> dml::Result<Vec<u8>> {
    snapshot.replace_family_with_limit(family, dml::theme::family::part::MAX_XML_BYTES)
}

#[cfg(feature = "drawingml-theme-family-current")]
fn theme_family_replace(
    snapshot: &dml::theme::family::part::Snapshot,
    family: &dml::theme::family::Family,
) -> dml::Result<Vec<u8>> {
    snapshot.replace_family(family, dml::theme::family::part::Limits::DEFAULT)
}

#[cfg(not(feature = "drawingml-theme-family-current"))]
fn theme_family_add(source: &[u8], family: &dml::theme::family::Family) -> dml::Result<Vec<u8>> {
    dml::theme::family::part::add_family_with_limit(
        source,
        family,
        dml::theme::family::part::MAX_XML_BYTES,
    )
}

#[cfg(feature = "drawingml-theme-family-current")]
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
        "<!--theme-part-before--><a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:thm15=\"{}\" name=\"{}\"><a:themeElements/><!--theme-part-opaque--><a:extLst><a:ext uri=\"{}\"> {family} </a:ext></a:extLst><!--theme-part-after--></a:theme>",
        dml::theme::family::NAMESPACE,
        THEME_FAMILY_FIXTURE_NAME,
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
    let result =
        dml::ink::actions::write_profile_to(writer, value, dml::ink::actions::Limits::DEFAULT);
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

fn read_geometry_value(bytes: &[u8]) -> Result<dml::geometry::CustomGeometry, ProbeFailure> {
    parse_geometry(bytes)
}

fn write_geometry_value(geometry: &dml::geometry::CustomGeometry) -> Result<Vec<u8>, ProbeFailure> {
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
    let result =
        dml::transform::codec::write_to(writer, value, &dml::transform::codec::Limits::DEFAULT);
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
        PartKind::Diagram => {
            let outline = read_diagram_value(bytes)?;
            let entry = outline
                .entries
                .first()
                .ok_or_else(|| fail("diagram data model has no content entry"))?;
            PreparedValue::Diagram {
                source: bytes.to_vec(),
                entry_id: entry.id.clone(),
                entry_text: entry.text.clone(),
                drawing_relationship: outline.drawing_relationship.clone(),
            }
        },
        PartKind::TextBody => PreparedValue::TextBody(read_text_value(bytes)?),
        PartKind::SvgBlip => PreparedValue::SvgBlip(read_svg_value(bytes)?),
        PartKind::Model3d => PreparedValue::Model3d(read_model3d_value(bytes)?),
        PartKind::Blip => PreparedValue::Blip(read_blip_value(bytes)?),
        PartKind::Theme => PreparedValue::Theme(read_theme_value(bytes)?),
        PartKind::ThemeFamily => PreparedValue::ThemeFamily(read_family_value(bytes)?),
        PartKind::ThemeFamilyPart => {
            let snapshot = read_theme_family_part_value(bytes)?;
            let family = snapshot
                .family()
                .cloned()
                .ok_or_else(|| fail("theme family part has no direct family owner"))?;
            PreparedValue::ThemeFamilyPart { snapshot, family }
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
        (
            PartKind::Diagram,
            PreparedValue::Diagram {
                source,
                entry_id,
                entry_text,
                drawing_relationship,
                ..
            },
        ) => {
            let output = set_diagram_text(source, entry_id, entry_text)?;
            let output = set_diagram_reference(&output, drawing_relationship.as_deref())?;
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
            #[cfg(feature = "drawingml-raw-current")]
            let result =
                dml::theme::family::write_to(writer, value, dml::theme::family::Limits::DEFAULT);
            #[cfg(not(feature = "drawingml-raw-current"))]
            let result = dml::theme::family::write_to(writer, value);
            result.map_err(|error| classify(error, 0))
        },
        (PartKind::ThemeFamilyPart, PreparedValue::ThemeFamilyPart { snapshot, family }) => {
            let output = theme_family_replace(snapshot, family)
                .map_err(|error| classify(error, snapshot.xml_bytes().len()))?;
            writer
                .write_all(&output)
                .map_err(|error| io_failure(error, output.len()))
        },
        (PartKind::InkActions, PreparedValue::InkActions(value)) => write_ink_value(writer, value),
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

fn write_prepared_vec(kind: PartKind, value: &PreparedValue) -> Result<Vec<u8>, ProbeFailure> {
    let mut output = Vec::new();
    write_prepared_to(kind, value, &mut output)?;
    Ok(output)
}

/// Output admission is a separate capability from a caller-provided sink.
/// Several DrawingML codecs only expose a Vec/String writer, but still take a
/// typed output budget and must be exercised at that boundary.  SVG retains
/// its historical bounded writer in the baseline; the other current raw
/// limits are enabled by their respective source fences.
fn output_limit_supported(kind: PartKind) -> bool {
    match kind {
        PartKind::SvgBlip => true,
        PartKind::Chart
        | PartKind::FormatCode
        | PartKind::Diagram
        | PartKind::TextBody
        | PartKind::Model3d
        | PartKind::Theme
        | PartKind::Geometry
        | PartKind::Color
        | PartKind::Transform => cfg!(feature = "drawingml-raw-current"),
        PartKind::ThemeFamily => cfg!(feature = "drawingml-raw-current"),
        PartKind::ThemeFamilyPart => cfg!(feature = "drawingml-theme-family-current"),
        PartKind::InkActions => cfg!(any(
            feature = "drawingml-raw-current",
            feature = "drawingml-ink-limits-current"
        )),
        PartKind::Blip => false,
    }
}

fn text_body_payload_bytes(bytes: &[u8]) -> Result<usize, ProbeFailure> {
    const PREFIX: &[u8] =
        b"<a:txBody xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">";
    const SUFFIX: &[u8] = b"</a:txBody>";
    let payload = bytes
        .strip_prefix(PREFIX)
        .and_then(|bytes| bytes.strip_suffix(SUFFIX))
        .ok_or_else(|| fail("text-body fixture is not wrapped by the probe host"))?;
    Ok(payload.len())
}

fn output_payload_bytes(part: &Part, prepared: &PreparedPart) -> Result<usize, ProbeFailure> {
    match part.kind {
        PartKind::TextBody => text_body_payload_bytes(&part.changed),
        PartKind::SvgBlip => prepared
            .contextual_changed_output_bytes
            .ok_or_else(|| fail("SVG contextual output was not prepared before timing")),
        _ => Ok(part.changed.len()),
    }
}

#[cfg(feature = "drawingml-raw-current")]
fn write_chart_value_limited<W: Write>(
    writer: &mut W,
    value: &dml::chart::Chart,
    maximum: usize,
) -> Result<(), ProbeFailure> {
    dml::chart::writer::write(writer, value, &dml::chart::Limits::new(maximum))
        .map_err(|error| classify(error, maximum.saturating_add(1)))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn write_chart_value_limited<W: Write>(
    _writer: &mut W,
    _value: &dml::chart::Chart,
    _maximum: usize,
) -> Result<(), ProbeFailure> {
    Err(fail(
        "chart output limits are unavailable in the historical API",
    ))
}

#[cfg(feature = "drawingml-raw-current")]
fn write_format_code_value_limited<W: Write>(
    writer: &mut W,
    value: &dml::chart::extension::formatcode2::Element,
    maximum: usize,
) -> Result<(), ProbeFailure> {
    let defaults = dml::chart::extension::formatcode2::Limits::DEFAULT;
    let limits = dml::chart::extension::formatcode2::Limits::new(
        defaults.max_xml_bytes,
        defaults.max_value_bytes,
        defaults.max_attributes,
        defaults.max_nodes,
        defaults.max_depth,
        defaults.max_namespace_bytes,
        maximum,
    );
    dml::chart::extension::formatcode2::write_to(writer, value, &limits)
        .map_err(|error| classify(error, maximum.saturating_add(1)))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn write_format_code_value_limited<W: Write>(
    _writer: &mut W,
    _value: &dml::chart::extension::formatcode2::Element,
    _maximum: usize,
) -> Result<(), ProbeFailure> {
    Err(fail(
        "formatcode2 output limits are unavailable in the historical API",
    ))
}

#[cfg(feature = "drawingml-raw-current")]
fn diagram_limits_with_output(maximum: usize) -> dml::diagram::Limits {
    let defaults = dml::diagram::Limits::DEFAULT;
    dml::diagram::Limits::new(
        defaults.input_bytes_value(),
        maximum,
        defaults.xml_nodes_value(),
        defaults.xml_depth_value(),
        defaults.points_value(),
        defaults.connections_value(),
        defaults.data_text_bytes_value(),
        defaults.nodes_value(),
        defaults.node_depth_value(),
        defaults.node_text_bytes_value(),
    )
}

#[cfg(feature = "drawingml-raw-current")]
fn write_diagram_value_limited(
    source: &[u8],
    entry_id: &str,
    entry_text: &str,
    drawing_relationship: Option<&str>,
    maximum: usize,
) -> Result<Vec<u8>, ProbeFailure> {
    let limits = diagram_limits_with_output(maximum);
    let output = dml::diagram::edit::set_text(source, entry_id, entry_text, &limits)
        .map_err(|error| classify(error, maximum.saturating_add(1)))?;
    dml::diagram::edit::set_drawing_reference(&output, drawing_relationship, &limits)
        .map_err(|error| classify(error, maximum.saturating_add(1)))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn write_diagram_value_limited(
    _source: &[u8],
    _entry_id: &str,
    _entry_text: &str,
    _drawing_relationship: Option<&str>,
    _maximum: usize,
) -> Result<Vec<u8>, ProbeFailure> {
    Err(fail(
        "diagram output limits are unavailable in the historical API",
    ))
}

#[cfg(feature = "drawingml-raw-current")]
fn write_text_value_limited(
    value: &dml::text::body::Body,
    maximum: usize,
) -> Result<String, ProbeFailure> {
    let mut output = String::new();
    dml::text::body::writer::write(
        &mut output,
        value,
        dml::text::body::writer::Limits::new(maximum),
    )
    .map_err(|error| classify(error, maximum.saturating_add(1)))?;
    Ok(output)
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn write_text_value_limited(
    _value: &dml::text::body::Body,
    _maximum: usize,
) -> Result<String, ProbeFailure> {
    Err(fail(
        "text-body output limits are unavailable in the historical API",
    ))
}

#[cfg(feature = "drawingml-raw-current")]
fn write_model3d_value_limited<W: Write>(
    writer: &mut W,
    value: &dml::model3d::Metadata,
    maximum: usize,
) -> Result<(), ProbeFailure> {
    let defaults = dml::model3d::codec::Limits::DEFAULT;
    dml::model3d::codec::write_to(
        writer,
        value,
        dml::model3d::codec::Limits::new(
            defaults.max_xml_bytes,
            defaults.max_fragment_bytes,
            maximum,
        ),
    )
    .map_err(|error| classify(error, maximum.saturating_add(1)))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn write_model3d_value_limited<W: Write>(
    _writer: &mut W,
    _value: &dml::model3d::Metadata,
    _maximum: usize,
) -> Result<(), ProbeFailure> {
    Err(fail(
        "model3d output limits are unavailable in the historical API",
    ))
}

#[cfg(feature = "drawingml-raw-current")]
fn write_theme_value_limited(
    value: &dml::theme::Theme,
    maximum: usize,
) -> Result<Vec<u8>, ProbeFailure> {
    dml::theme::codec::encode_part(
        &value.name,
        &value.colors,
        &value.fonts,
        dml::theme::codec::Limits::new(maximum),
    )
    .map_err(|error| classify(error, maximum.saturating_add(1)))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn write_theme_value_limited(
    _value: &dml::theme::Theme,
    _maximum: usize,
) -> Result<Vec<u8>, ProbeFailure> {
    Err(fail(
        "theme output limits are unavailable in the historical API",
    ))
}

#[cfg(feature = "drawingml-raw-current")]
fn write_family_value_limited<W: Write>(
    writer: &mut W,
    value: &dml::theme::family::Family,
    maximum: usize,
) -> Result<(), ProbeFailure> {
    dml::theme::family::write_to(writer, value, dml::theme::family::Limits::new(maximum))
        .map_err(|error| classify(error, maximum.saturating_add(1)))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn write_family_value_limited<W: Write>(
    _writer: &mut W,
    _value: &dml::theme::family::Family,
    _maximum: usize,
) -> Result<(), ProbeFailure> {
    Err(fail(
        "theme-family output limits are unavailable in the historical API",
    ))
}

#[cfg(feature = "drawingml-theme-family-current")]
fn write_theme_family_part_limited(
    snapshot: &dml::theme::family::part::Snapshot,
    family: &dml::theme::family::Family,
    maximum: usize,
) -> Result<Vec<u8>, ProbeFailure> {
    snapshot
        .replace_family(family, dml::theme::family::part::Limits::new(maximum))
        .map_err(|error| classify(error, maximum.saturating_add(1)))
}

#[cfg(not(feature = "drawingml-theme-family-current"))]
fn write_theme_family_part_limited(
    _snapshot: &dml::theme::family::part::Snapshot,
    _family: &dml::theme::family::Family,
    _maximum: usize,
) -> Result<Vec<u8>, ProbeFailure> {
    Err(fail(
        "theme-family-part output limits are unavailable in the historical API",
    ))
}

#[cfg(any(
    feature = "drawingml-raw-current",
    feature = "drawingml-ink-limits-current"
))]
fn write_ink_value_limited<W: Write>(
    writer: &mut W,
    value: &dml::ink::actions::Profile,
    maximum: usize,
) -> Result<(), ProbeFailure> {
    let defaults = dml::ink::actions::Limits::DEFAULT;
    dml::ink::actions::write_profile_to(
        writer,
        value,
        dml::ink::actions::Limits::new(defaults.max_source_bytes, maximum),
    )
    .map_err(|error| classify(error, maximum.saturating_add(1)))
}

#[cfg(not(any(
    feature = "drawingml-raw-current",
    feature = "drawingml-ink-limits-current"
)))]
fn write_ink_value_limited<W: Write>(
    _writer: &mut W,
    _value: &dml::ink::actions::Profile,
    _maximum: usize,
) -> Result<(), ProbeFailure> {
    Err(fail(
        "InkML output limits are unavailable in the historical API",
    ))
}

#[cfg(feature = "drawingml-raw-current")]
fn write_geometry_value_limited(
    value: &dml::geometry::CustomGeometry,
    maximum: usize,
) -> Result<String, ProbeFailure> {
    let mut output = String::new();
    let mut limits = dml::geometry::Limits::DEFAULT;
    limits.max_output_bytes = maximum;
    dml::geometry::writer::write_custom_geometry(&mut output, value, &limits)
        .map_err(|error| classify(error, maximum.saturating_add(1)))?;
    Ok(output)
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn write_geometry_value_limited(
    _value: &dml::geometry::CustomGeometry,
    _maximum: usize,
) -> Result<String, ProbeFailure> {
    Err(fail(
        "geometry output limits are unavailable in the historical API",
    ))
}

#[cfg(feature = "drawingml-raw-current")]
fn write_color_value_limited(
    value: &dml::color::Value,
    maximum: usize,
) -> Result<Vec<u8>, ProbeFailure> {
    let defaults = dml::color::Limits::DEFAULT;
    dml::color::write(
        value,
        &dml::color::Limits::new(
            maximum,
            defaults.max_depth,
            defaults.max_nodes,
            defaults.max_transforms,
        ),
    )
    .map_err(|error| classify(error, maximum.saturating_add(1)))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn write_color_value_limited(
    _value: &dml::color::Value,
    _maximum: usize,
) -> Result<Vec<u8>, ProbeFailure> {
    Err(fail(
        "color output limits are unavailable in the historical API",
    ))
}

#[cfg(feature = "drawingml-raw-current")]
fn write_transform_value_limited<W: Write>(
    writer: &mut W,
    value: &dml::transform::Transform,
    maximum: usize,
) -> Result<(), ProbeFailure> {
    let defaults = dml::transform::codec::Limits::DEFAULT;
    dml::transform::codec::write_to(
        writer,
        value,
        &dml::transform::codec::Limits::new(maximum, defaults.max_nodes),
    )
    .map_err(|error| classify(error, maximum.saturating_add(1)))
}

#[cfg(not(feature = "drawingml-raw-current"))]
fn write_transform_value_limited<W: Write>(
    _writer: &mut W,
    _value: &dml::transform::Transform,
    _maximum: usize,
) -> Result<(), ProbeFailure> {
    Err(fail(
        "transform output limits are unavailable in the historical API",
    ))
}

fn output_limit_result(
    part: &Part,
    prepared: &PreparedPart,
    maximum: usize,
    mut sink: &mut CountingSink,
) -> Result<(), ProbeFailure> {
    let result = match (part.kind, &prepared.changed) {
        (PartKind::Chart, PreparedValue::Chart(value)) => {
            write_chart_value_limited(&mut sink, value, maximum)
        },
        (PartKind::FormatCode, PreparedValue::FormatCode(value)) => {
            write_format_code_value_limited(&mut sink, value, maximum)
        },
        (
            PartKind::Diagram,
            PreparedValue::Diagram {
                source,
                entry_id,
                entry_text,
                drawing_relationship,
                ..
            },
        ) => write_diagram_value_limited(
            source,
            entry_id,
            entry_text,
            drawing_relationship.as_deref(),
            maximum,
        )
        .map(|_| ()),
        (PartKind::TextBody, PreparedValue::TextBody(value)) => {
            write_text_value_limited(value, maximum).map(|_| ())
        },
        (PartKind::SvgBlip, _) => {
            let value = prepared
                .contextual_changed
                .as_ref()
                .ok_or_else(|| fail("SVG contextual value was not prepared"))?;
            write_svg_contextual_limited(&mut sink, value, maximum)
        },
        (PartKind::Model3d, PreparedValue::Model3d(value)) => {
            write_model3d_value_limited(&mut sink, value, maximum)
        },
        (PartKind::Theme, PreparedValue::Theme(value)) => {
            write_theme_value_limited(value, maximum).map(|_| ())
        },
        (PartKind::ThemeFamily, PreparedValue::ThemeFamily(value)) => {
            write_family_value_limited(&mut sink, value, maximum)
        },
        (PartKind::ThemeFamilyPart, PreparedValue::ThemeFamilyPart { snapshot, family }) => {
            write_theme_family_part_limited(snapshot, family, maximum).map(|_| ())
        },
        (PartKind::InkActions, PreparedValue::InkActions(value)) => {
            write_ink_value_limited(&mut sink, value, maximum)
        },
        (PartKind::Geometry, PreparedValue::Geometry(value)) => {
            write_geometry_value_limited(value, maximum).map(|_| ())
        },
        (PartKind::Color, PreparedValue::Color(value)) => {
            write_color_value_limited(value, maximum).map(|_| ())
        },
        (PartKind::Transform, PreparedValue::Transform(value)) => {
            write_transform_value_limited(&mut sink, value, maximum)
        },
        (PartKind::Blip, _) => return Err(fail("image blip has no typed output writer")),
        _ => return Err(fail("prepared raw value kind mismatch")),
    };
    result
}

fn output_limit_attempt(
    part: &Part,
    prepared: &PreparedPart,
    maximum: usize,
) -> Result<OutputLimitAttempt, ProbeFailure> {
    let mut sink = CountingSink::new(usize::MAX);
    let result = output_limit_result(part, prepared, maximum, &mut sink);
    let error =
        result.expect_err("typed output-cap preflight unexpectedly accepted the selected cap");
    let refusal = error
        .refusal
        .ok_or_else(|| fail(format!("{} output cap lacked a typed refusal", part.id)))?;
    if refusal.kind != "limit" {
        return Err(fail(format!(
            "{} output cap returned {} instead of limit",
            part.id, refusal.kind
        )));
    }
    let partial_output_allowed = part.kind == PartKind::Chart;
    if partial_output_allowed {
        if sink.bytes > maximum {
            return Err(fail(format!(
                "{} bounded partial output exceeded its cap",
                part.id
            )));
        }
    } else if sink.writes != 0 || sink.bytes != 0 {
        return Err(fail(format!(
            "{} output preflight wrote before refusal",
            part.id
        )));
    }
    let sink = if supports_production_sink(part.kind) {
        Some(write_receipt(sink))
    } else {
        None
    };
    Ok(OutputLimitAttempt {
        refusal,
        sink,
        partial_output_allowed,
    })
}

/// Timed output refusal calls deliberately keep the concrete production
/// `Result` opaque.  This helper has no common probe error, refusal mapping,
/// digest, or receipt path; those belong exclusively to `output_limit_attempt`
/// in the pre-warmup detail pass.
fn output_limit_timed(
    part: &Part,
    prepared: &PreparedPart,
    maximum: usize,
    sink: &mut CountingSink,
) {
    match (part.kind, &prepared.changed) {
        (PartKind::Chart, PreparedValue::Chart(value)) => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::chart::writer::write(
                sink,
                value,
                &dml::chart::Limits::new(maximum),
            ));
        },
        (PartKind::FormatCode, PreparedValue::FormatCode(value)) => {
            #[cfg(feature = "drawingml-raw-current")]
            {
                let defaults = dml::chart::extension::formatcode2::Limits::DEFAULT;
                let limits = dml::chart::extension::formatcode2::Limits::new(
                    defaults.max_xml_bytes,
                    defaults.max_value_bytes,
                    defaults.max_attributes,
                    defaults.max_nodes,
                    defaults.max_depth,
                    defaults.max_namespace_bytes,
                    maximum,
                );
                black_box(dml::chart::extension::formatcode2::write_to(
                    sink, value, &limits,
                ));
            }
        },
        (
            PartKind::Diagram,
            PreparedValue::Diagram {
                source,
                entry_id,
                entry_text,
                drawing_relationship,
                ..
            },
        ) => {
            #[cfg(feature = "drawingml-raw-current")]
            {
                let limits = diagram_limits_with_output(maximum);
                let result = dml::diagram::edit::set_text(source, entry_id, entry_text, &limits);
                let result = result.and_then(|output| {
                    dml::diagram::edit::set_drawing_reference(
                        &output,
                        drawing_relationship.as_deref(),
                        &limits,
                    )
                });
                black_box(result);
            }
            #[cfg(not(feature = "drawingml-raw-current"))]
            {
                let result = dml::diagram::edit::set_text(source, entry_id, entry_text);
                let result = result.and_then(|output| {
                    dml::diagram::edit::set_drawing_reference(
                        &output,
                        drawing_relationship.as_deref(),
                    )
                });
                black_box(result);
            }
        },
        (PartKind::TextBody, PreparedValue::TextBody(value)) => {
            #[cfg(feature = "drawingml-raw-current")]
            {
                let mut output = String::new();
                black_box(dml::text::body::writer::write(
                    &mut output,
                    value,
                    dml::text::body::writer::Limits::new(maximum),
                ));
            }
        },
        (PartKind::SvgBlip, _) => {
            if let Some(value) = prepared.contextual_changed.as_ref() {
                #[cfg(feature = "drawingml-raw-current")]
                {
                    let defaults = dml::svg_blip::Limits::DEFAULT;
                    let limits = dml::svg_blip::Limits::new(defaults.max_input_bytes, maximum);
                    black_box(dml::svg_blip::write_contextual_to(sink, value, limits));
                }
                #[cfg(not(feature = "drawingml-raw-current"))]
                black_box(dml::svg_blip::write_contextual_to(sink, value, maximum));
            }
        },
        (PartKind::Model3d, PreparedValue::Model3d(value)) => {
            #[cfg(feature = "drawingml-raw-current")]
            {
                let defaults = dml::model3d::codec::Limits::DEFAULT;
                let limits = dml::model3d::codec::Limits::new(
                    defaults.max_xml_bytes,
                    defaults.max_fragment_bytes,
                    maximum,
                );
                black_box(dml::model3d::codec::write_to(sink, value, limits));
            }
        },
        (PartKind::Theme, PreparedValue::Theme(value)) => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::theme::codec::encode_part(
                &value.name,
                &value.colors,
                &value.fonts,
                dml::theme::codec::Limits::new(maximum),
            ));
        },
        (PartKind::ThemeFamily, PreparedValue::ThemeFamily(value)) => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::theme::family::write_to(
                sink,
                value,
                dml::theme::family::Limits::new(maximum),
            ));
        },
        (PartKind::ThemeFamilyPart, PreparedValue::ThemeFamilyPart { snapshot, family }) => {
            #[cfg(feature = "drawingml-theme-family-current")]
            black_box(
                snapshot.replace_family(family, dml::theme::family::part::Limits::new(maximum)),
            );
        },
        (PartKind::InkActions, PreparedValue::InkActions(value)) => {
            #[cfg(any(
                feature = "drawingml-raw-current",
                feature = "drawingml-ink-limits-current"
            ))]
            {
                let defaults = dml::ink::actions::Limits::DEFAULT;
                black_box(dml::ink::actions::write_profile_to(
                    sink,
                    value,
                    dml::ink::actions::Limits::new(defaults.max_source_bytes, maximum),
                ));
            }
        },
        (PartKind::Geometry, PreparedValue::Geometry(value)) => {
            #[cfg(feature = "drawingml-raw-current")]
            {
                let mut output = String::new();
                let mut limits = dml::geometry::Limits::DEFAULT;
                limits.max_output_bytes = maximum;
                black_box(dml::geometry::writer::write_custom_geometry(
                    &mut output,
                    value,
                    &limits,
                ));
            }
        },
        (PartKind::Color, PreparedValue::Color(value)) => {
            #[cfg(feature = "drawingml-raw-current")]
            {
                let defaults = dml::color::Limits::DEFAULT;
                black_box(dml::color::write(
                    value,
                    &dml::color::Limits::new(
                        maximum,
                        defaults.max_depth,
                        defaults.max_nodes,
                        defaults.max_transforms,
                    ),
                ));
            }
        },
        (PartKind::Transform, PreparedValue::Transform(value)) => {
            #[cfg(feature = "drawingml-raw-current")]
            {
                let defaults = dml::transform::codec::Limits::DEFAULT;
                black_box(dml::transform::codec::write_to(
                    sink,
                    value,
                    &dml::transform::codec::Limits::new(maximum, defaults.max_nodes),
                ));
            }
        },
        _ => {},
    }
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
        sha256: hex_digest(
            sink.digest
                .expect("timed sink cannot produce a receipt")
                .finalize(),
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

fn build_generated_fixture_parts(case: &LaneCase) -> Vec<GeneratedFixturePart> {
    case.parts
        .iter()
        .map(|part| GeneratedFixturePart {
            part_id: part.id,
            source_sha256: sha256(&part.source),
            source_bytes: part.source.len(),
            changed_sha256: sha256(&part.changed),
            changed_bytes: part.changed.len(),
            malformed_sha256: sha256(&part.malformed),
            malformed_bytes: part.malformed.len(),
        })
        .collect()
}

fn generated_fixture_identity(parts: &[GeneratedFixturePart]) -> String {
    let mut identity = String::from("drawingml-generated-fixture-v1\0");
    for part in parts {
        identity.push_str(part.part_id);
        identity.push('\0');
        identity.push_str(&part.source_sha256);
        identity.push('\0');
        identity.push_str(&part.source_bytes.to_string());
        identity.push('\0');
        identity.push_str(&part.changed_sha256);
        identity.push('\0');
        identity.push_str(&part.changed_bytes.to_string());
        identity.push('\0');
        identity.push_str(&part.malformed_sha256);
        identity.push('\0');
        identity.push_str(&part.malformed_bytes.to_string());
        identity.push('\0');
    }
    sha256(identity.as_bytes())
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
    let over = over_limit(PartKind::Chart, &source);
    Ok(vec![Part {
        id: "chart",
        kind: PartKind::Chart,
        source,
        changed,
        malformed: malformed_fragment("chart", &fixture.malformed),
        // chart::reader::read delegates its input admission to the shared
        // MCE default policy (256 MiB).  Keep the boundary fixture tied to
        // that real production limit rather than treating Chart as unbounded.
        over,
        limit_resource: "MCE input bytes",
    }])
}

fn format_code_part(fixture: &FixtureLane) -> Result<Part, ProbeFailure> {
    // Keep a deliberately non-canonical source (comments, surrounding
    // whitespace, and CDATA) so the changed fixture proves the production
    // source splice rather than a detached canonical round trip.
    let detached = dml::chart::extension::formatcode2::Element::new("0.00")
        .map_err(|error| classify(error, 0))?;
    let canonical = {
        let mut output = Vec::new();
        write_format_code_value(&mut output, &detached)?;
        output
    };
    let canonical_text = std::str::from_utf8(&canonical)
        .map_err(|error| invalid_input(error.to_string(), canonical.len()))?;
    let open_end = canonical_text
        .find('>')
        .ok_or_else(|| fail("formatcode2 detached fixture has no root opening tag"))?;
    let root_open = &canonical_text[..=open_end];
    let source = format!(
        "<!--formatcode2-before-->\n{root_open}<![CDATA[0.00]]><!--formatcode2-opaque--></c16r2:formatcode2>\n<!--formatcode2-after-->"
    )
    .into_bytes();
    let mut value = read_format_code_value(&source)?;
    value
        .set_value(format!("0.00;{}", fixture.mutation.len()))
        .map_err(|error| classify(error, 0))?;
    let changed = {
        let mut output = Vec::new();
        write_format_code_value(&mut output, &value)?;
        output
    };
    let over = over_limit(PartKind::FormatCode, &source);
    Ok(Part {
        id: "formatcode2",
        kind: PartKind::FormatCode,
        source,
        changed,
        malformed: malformed_fragment("formatcode2", &fixture.malformed),
        over,
        limit_resource: "formatcode2 XML bytes",
    })
}

fn diagram_parts(fixture: &FixtureLane) -> Result<Vec<Part>, ProbeFailure> {
    use dml::diagram::{DiagramType, SmartArtBuilder, generate_smartart_data_xml};
    let diagram = SmartArtBuilder::new(DiagramType::Relationship)
        .layout_name("relationship")
        .add_items([fixture_seed(fixture), "second".to_string()])
        .build();
    #[cfg(feature = "drawingml-raw-current")]
    let source = generate_smartart_data_xml(&diagram, &dml::diagram::Limits::DEFAULT)
        .map_err(|error| classify(error, 0))?
        .into_bytes();
    #[cfg(not(feature = "drawingml-raw-current"))]
    let source = generate_smartart_data_xml(&diagram).into_bytes();
    let outline = read_diagram_value(&source)?;
    let entry = outline
        .entries
        .first()
        .ok_or_else(|| fail("diagram authoring produced no content entry"))?;
    let text_changed =
        set_diagram_text(&source, &entry.id, &format!("{} changed", fixture.mutation))?;
    let changed = set_diagram_reference(&text_changed, Some("rIdDrawing"))?;
    let over = over_limit(PartKind::Diagram, &source);
    Ok(vec![Part {
        id: "diagram-data-model",
        kind: PartKind::Diagram,
        source,
        changed,
        malformed: malformed_fragment("diagram-data-model", &fixture.malformed),
        over,
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
    // This is intentionally a host fragment: all four prefixes are inherited
    // from the surrounding part and the root carries no xmlns declarations.
    // The contextual reader below is therefore required for both the source
    // and changed fixtures; a self-declaring root would not exercise the
    // namespace-scope API.
    let svg_source = br#"<asvg:svgBlip r:embed="rIdSvg" future:marker="opaque:value"><opaque:payload opaque:kind="future"/></asvg:svgBlip>"#.to_vec();
    let svg_context = svg_fixture_context()?;
    let mut svg = read_svg_contextual_value_with(&svg_source, &svg_context)?;
    svg.set_reference(
        dml::svg_blip::Reference::embedded("rIdSvgChanged")
            .map_err(|error| fail(error.to_string()))?,
    )
    .map_err(|error| classify(error, svg_source.len()))?;
    let svg_changed = write_svg_contextual_default(&svg)?;
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
    let text_over = over_limit(PartKind::TextBody, &source);
    let svg_over = over_limit(PartKind::SvgBlip, &svg_source);
    let model_over = over_limit(PartKind::Model3d, &model_source);
    Ok(vec![
        Part {
            id: "text-body",
            kind: PartKind::TextBody,
            source,
            changed,
            malformed: malformed_fragment("text-body", &fixture.malformed),
            over: text_over,
            limit_resource: "text body input bytes",
        },
        Part {
            id: "svg-blip",
            kind: PartKind::SvgBlip,
            source: svg_source,
            changed: svg_changed,
            malformed: malformed_fragment("svg-blip", &fixture.malformed),
            over: svg_over,
            limit_resource: "SVG blip XML bytes",
        },
        Part {
            id: "model3d",
            kind: PartKind::Model3d,
            source: model_source,
            changed: model_changed,
            malformed: malformed_fragment("model3d", &fixture.malformed),
            over: model_over,
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

const THEME_FAMILY_FIXTURE_NAME: &str = "Office";
const THEME_FAMILY_FIXTURE_ID: &str = "{11111111-1111-1111-1111-111111111111}";
const THEME_FAMILY_FIXTURE_VARIANT_ID: &str = "{22222222-2222-2222-2222-222222222222}";

fn theme_family_fixture_identity() -> (&'static str, &'static str, &'static str) {
    (
        THEME_FAMILY_FIXTURE_NAME,
        THEME_FAMILY_FIXTURE_ID,
        THEME_FAMILY_FIXTURE_VARIANT_ID,
    )
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
        THEME_FAMILY_FIXTURE_NAME,
        THEME_FAMILY_FIXTURE_ID,
        THEME_FAMILY_FIXTURE_VARIANT_ID,
    )
    .map_err(|error| classify(error, 0))?;
    let family_canonical = write_family_value(&family)?;
    let family_canonical = std::str::from_utf8(&family_canonical)
        .map_err(|error| invalid_input(error.to_string(), family_canonical.len()))?;
    let family_source =
        format!("<!--theme-family-before-->\n  {family_canonical}\n<!--theme-family-after-->")
            .into_bytes();
    // Parse the non-canonical source, then mutate the retained projection.
    // The changed bytes therefore exercise the production lexical splice.
    let mut family_value = read_family_value(&family_source)?;
    family_value
        .set_name(format!("Family {}", fixture.mutation.len()))
        .map_err(|error| classify(error, family_source.len()))?;
    let family_changed = write_family_value(&family_value)?;
    let family_part_source = complete_theme_family_part(&family_source)?;
    // Mutate the complete source-backed part rather than embedding the
    // detached family serialization into a newly generated wrapper.
    let family_part_snapshot = read_theme_family_part_value(&family_part_source)?;
    let source_family = family_part_snapshot
        .family()
        .cloned()
        .ok_or_else(|| fail("complete Theme fixture has no direct family owner"))?;
    let (expected_name, expected_id, expected_variant_id) = theme_family_fixture_identity();
    if source_family.name() != expected_name
        || source_family.id().to_string() != expected_id
        || source_family.variant_id().to_string() != expected_variant_id
    {
        return Err(fail(
            "complete Theme fixture source identity does not match its bound name/id/variant",
        ));
    }
    let family_part_changed = theme_family_replace(&family_part_snapshot, &family_value)
        .map_err(|error| classify(error, family_part_source.len()))?;
    if source_family == family_value {
        return Err(fail(
            "Theme fixture mutation did not change family identity",
        ));
    }
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
    // The writer is intentionally exercised by the raw fidelity route; do
    // not use its canonical bytes as the source fixture for this lane.
    let _ = ink_profile;
    let theme_over = over_limit(PartKind::Theme, &source);
    let family_over = over_limit(PartKind::ThemeFamily, &family_source);
    let family_part_over = over_limit(PartKind::ThemeFamilyPart, &family_part_source);
    let ink_over = over_limit(PartKind::InkActions, &ink_source);
    Ok(vec![
        Part {
            id: "theme",
            kind: PartKind::Theme,
            source,
            changed,
            malformed: malformed_fragment("theme", &fixture.malformed),
            over: theme_over,
            limit_resource: "theme XML bytes",
        },
        Part {
            id: "theme-family",
            kind: PartKind::ThemeFamily,
            source: family_source,
            changed: family_changed,
            malformed: malformed_fragment("theme-family", &fixture.malformed),
            over: family_over,
            limit_resource: "theme family XML bytes",
        },
        Part {
            id: "theme-family-part",
            kind: PartKind::ThemeFamilyPart,
            source: family_part_source,
            changed: family_part_changed,
            malformed: malformed_fragment("theme-family-part", &fixture.malformed),
            over: family_part_over,
            limit_resource: "Theme Family XML bytes",
        },
        Part {
            id: "ink-actions",
            kind: PartKind::InkActions,
            source: ink_source,
            changed: ink_changed,
            malformed: malformed_fragment("ink-actions", &fixture.malformed),
            over: ink_over,
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
    let color_source =
        write_color_value(&dml::color::Value::rgb(dml::color::Rgb::new(51, 102, 153)))?;
    let color_changed =
        write_color_value(&dml::color::Value::rgb(dml::color::Rgb::new(153, 102, 51)))?;
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
    let color_over = over_limit(PartKind::Color, &color_source);
    let transform_over = over_limit(PartKind::Transform, &transform_source);
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
            over: color_over,
            limit_resource: "DrawingML color XML",
        },
        format_part,
        Part {
            id: "transform",
            kind: PartKind::Transform,
            source: transform_source,
            changed: transform_changed,
            malformed: malformed_fragment("transform", &fixture.malformed),
            over: transform_over,
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
    let mut prepared = parts
        .iter()
        .map(|part| {
            let contextual_namespace = if part.kind == PartKind::SvgBlip {
                Some(svg_fixture_context()?)
            } else {
                None
            };
            let contextual_source = if part.kind == PartKind::SvgBlip {
                Some(read_svg_contextual_value_with(
                    &part.source,
                    contextual_namespace
                        .as_ref()
                        .ok_or_else(|| fail("SVG namespace context was not prepared"))?,
                )?)
            } else {
                None
            };
            let contextual_changed = if part.kind == PartKind::SvgBlip {
                Some(read_svg_contextual_value_with(
                    &part.changed,
                    contextual_namespace
                        .as_ref()
                        .ok_or_else(|| fail("SVG namespace context was not prepared"))?,
                )?)
            } else {
                None
            };
            let contextual_source_output_bytes = contextual_source
                .as_ref()
                .map(write_svg_contextual_default)
                .transpose()?
                .map(|output| output.len());
            let contextual_changed_output_bytes = contextual_changed
                .as_ref()
                .map(write_svg_contextual_default)
                .transpose()?
                .map(|output| output.len());
            Ok(PreparedPart {
                id: part.id,
                source: prepare_value(part.kind, &part.source)?,
                changed: prepare_value(part.kind, &part.changed)?,
                contextual_source,
                contextual_changed,
                contextual_namespace,
                contextual_source_output_bytes,
                contextual_changed_output_bytes,
                output_payload_bytes: None,
                output_limit: None,
            })
        })
        .collect::<Result<Vec<_>, ProbeFailure>>()?;
    let shared: Vec<SharedPart> = parts
        .iter()
        .filter(|part| supports_shared(part.kind))
        .map(|part| SharedPart {
            id: part.id,
            kind: part.kind,
            source: std::sync::Arc::<[u8]>::from(part.source.clone()),
        })
        .collect();
    for (part, prepared_part) in parts.iter().zip(&mut prepared) {
        if output_limit_supported(part.kind) {
            let payload = output_payload_bytes(part, prepared_part)?;
            if payload <= 1 {
                return Err(fail(format!(
                    "{} output fixture is too small for a strict cap",
                    part.id
                )));
            }
            prepared_part.output_payload_bytes = Some(payload);
            prepared_part.output_limit = Some(payload - 1);
        }
    }
    let shared_prepared = shared
        .iter()
        .map(read_shared_value)
        .collect::<Result<Vec<_>, _>>()?;
    Ok(LaneCase {
        parts,
        prepared,
        shared,
        shared_prepared,
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
    let prepared = owner_prepare(fixture)?;
    // Build the changed patch from the same retained native owner that the
    // timed commit helper consumes.  In particular, current Ink actions must
    // use `Prepared::edit`, rather than reopening the source through the
    // detached legacy constructor merely to obtain a patch.
    let (changed, patch) =
        owner_native_commit_parts(route, owner_native_commit_prepared(route, &prepared, true))?;
    let inverse_patch = owner_patch_inverse(&patch);
    let mut changed_fixture = fixture.clone();
    changed_fixture.source = changed.clone();
    let changed_prepared = owner_prepare(&changed_fixture)?;
    let mut stale_fixture = fixture.clone();
    stale_fixture.source = stale_fixture.stale.clone();
    let stale_prepared = owner_prepare(&stale_fixture)?;
    case.prepared_owners.push(OwnerWork {
        fixture: fixture.clone(),
        changed,
        patch,
        inverse_patch,
        prepared,
        changed_prepared,
        stale_prepared,
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
    case.shared_prepared = case
        .shared
        .iter()
        .map(read_shared_value)
        .collect::<Result<Vec<_>, _>>()?;
    case.owners.clear();
    case.prepared_owners.clear();
    Ok(())
}

/// Build the inverse once, before a timed replay sample.  Inversion is part of
/// the proof/preparation work; constructing it in the sample would measure a
/// patch clone and its source-state swap in addition to the native replay.
fn owner_patch_inverse(patch: &OwnerPatch) -> OwnerPatch {
    match patch {
        OwnerPatch::DrawingmlChart(patch) => OwnerPatch::DrawingmlChart(patch.clone().inverse()),
        OwnerPatch::DrawingmlThemeFamily(patch) => {
            OwnerPatch::DrawingmlThemeFamily(patch.clone().inverse())
        },
        OwnerPatch::DrawingmlInkActions(patch) => OwnerPatch::DrawingmlInkActions(patch.inverse()),
        OwnerPatch::DrawingmlTransform(patch) => {
            OwnerPatch::DrawingmlTransform(patch.clone().inverse())
        },
    }
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
fn owner_prepare(fixture: &OwnerFixture) -> Result<OwnerPrepared, ProbeFailure> {
    match fixture.route {
        OwnerRoute::DrawingmlChart => dml::chart::edit::Snapshot::from_xml(fixture.source.clone())
            .map(OwnerPrepared::DrawingmlChart)
            .map_err(|error| owner_error(fixture.route, error)),
        OwnerRoute::DrawingmlThemeFamily => {
            dml::theme::family::Snapshot::from_xml(fixture.source.clone())
                .map(OwnerPrepared::DrawingmlThemeFamily)
                .map_err(|error| owner_error(fixture.route, error))
        },
        OwnerRoute::DrawingmlInkActions => dml::ink::actions::read_profile(&fixture.source)
            .map(OwnerPrepared::DrawingmlInkActions)
            .map_err(|error| owner_error(fixture.route, error)),
        OwnerRoute::DrawingmlTransform => {
            dml::transform::Snapshot::from_xml(fixture.source.clone())
                .map(OwnerPrepared::DrawingmlTransform)
                .map_err(|error| owner_error(fixture.route, error))
        },
    }
}

#[cfg(feature = "drawingml-owner-current")]
fn owner_prepare(fixture: &OwnerFixture) -> Result<OwnerPrepared, ProbeFailure> {
    match fixture.route {
        OwnerRoute::DrawingmlChart => dml::chart::edit::Package::from_xml(fixture.source.clone())
            .map(OwnerPrepared::DrawingmlChart)
            .map_err(|error| owner_error(fixture.route, error)),
        OwnerRoute::DrawingmlThemeFamily => owner_theme_package(fixture.source.clone())
            .map(OwnerPrepared::DrawingmlThemeFamily)
            .map_err(|error| owner_error(fixture.route, error)),
        OwnerRoute::DrawingmlInkActions => {
            use dml::ink::actions::{Edit, EditOptions};
            let profile = owner_read_profile(&fixture.source)
                .map_err(|error| owner_error(fixture.route, error))?;
            let prepared = Edit::from_profile(&profile, EditOptions::default())
                .map_err(|error| owner_error(fixture.route, error))?
                .commit()
                .map_err(|error| owner_error(fixture.route, error))?
                .into_package();
            Ok(OwnerPrepared::DrawingmlInkActions(prepared))
        },
        OwnerRoute::DrawingmlTransform => owner_transform_package(fixture.source.clone())
            .map(OwnerPrepared::DrawingmlTransform)
            .map_err(|error| owner_error(fixture.route, error)),
    }
}

#[cfg(not(feature = "drawingml-ink-limits-current"))]
fn owner_read_profile(source: &[u8]) -> dml::Result<dml::ink::actions::Profile> {
    dml::ink::actions::read_profile(source)
}

#[cfg(feature = "drawingml-ink-limits-current")]
fn owner_read_profile(source: &[u8]) -> dml::Result<dml::ink::actions::Profile> {
    dml::ink::actions::read_profile(source, dml::ink::actions::Limits::DEFAULT)
}

#[cfg(all(
    feature = "drawingml-owner-current",
    feature = "drawingml-theme-family-owner-limits-current"
))]
fn owner_theme_package(source: Vec<u8>) -> dml::Result<dml::theme::family::Package> {
    dml::theme::family::Package::from_xml(source, dml::theme::family::Limits::DEFAULT)
}

#[cfg(all(
    feature = "drawingml-owner-current",
    not(feature = "drawingml-theme-family-owner-limits-current")
))]
fn owner_theme_package(source: Vec<u8>) -> dml::Result<dml::theme::family::Package> {
    dml::theme::family::Package::from_xml(source)
}

#[cfg(all(
    feature = "drawingml-owner-current",
    feature = "drawingml-transform-limits-current"
))]
fn owner_transform_package(source: Vec<u8>) -> dml::Result<dml::transform::Package> {
    dml::transform::Package::from_xml(source, &dml::transform::Limits::DEFAULT)
}

#[cfg(all(
    feature = "drawingml-owner-current",
    not(feature = "drawingml-transform-limits-current")
))]
fn owner_transform_package(source: Vec<u8>) -> dml::Result<dml::transform::Package> {
    dml::transform::Package::from_xml(source)
}

#[cfg(not(feature = "drawingml-owner-current"))]
enum NativeOwnerResult {
    DrawingmlChart(dml::Result<dml::chart::edit::Commit>),
    DrawingmlThemeFamily(dml::Result<dml::theme::family::Commit>),
    DrawingmlInkActions(dml::Result<dml::ink::actions::Commit>),
    DrawingmlTransform(dml::Result<dml::transform::Commit>),
}

#[cfg(feature = "drawingml-owner-current")]
enum NativeOwnerResult {
    DrawingmlChart(dml::Result<dml::chart::edit::Commit>),
    DrawingmlThemeFamily(dml::Result<dml::theme::family::Commit>),
    DrawingmlInkActions(dml::Result<dml::ink::actions::Commit>),
    DrawingmlTransform(dml::Result<dml::transform::Commit>),
}

#[cfg(not(feature = "drawingml-owner-current"))]
enum NativeReplayResult {
    DrawingmlChart(dml::Result<dml::chart::edit::Snapshot>),
    DrawingmlThemeFamily(dml::Result<dml::theme::family::Snapshot>),
    DrawingmlInkActions(dml::Result<Vec<u8>>),
    DrawingmlTransform(dml::Result<dml::transform::Snapshot>),
}

#[cfg(feature = "drawingml-owner-current")]
enum NativeReplayResult {
    DrawingmlChart(dml::Result<dml::chart::edit::Commit>),
    DrawingmlThemeFamily(dml::Result<dml::theme::family::Commit>),
    DrawingmlInkActions(dml::Result<dml::ink::actions::Commit>),
    DrawingmlTransform(dml::Result<dml::transform::Commit>),
}

/// Run one owner edit against an already parsed/retained native owner.  This
/// helper deliberately returns the native result without converting it to
/// bytes or a probe error; the timed caller black-boxes that result directly.
#[cfg(not(feature = "drawingml-owner-current"))]
fn owner_native_commit_prepared(
    route: OwnerRoute,
    prepared: &OwnerPrepared,
    changed: bool,
) -> NativeOwnerResult {
    match (prepared, route) {
        (OwnerPrepared::DrawingmlChart(snapshot), OwnerRoute::DrawingmlChart) => {
            let result = (|| {
                let mut edit = snapshot.edit();
                if changed {
                    edit.set_title(Some(dml::chart::data::TitleText::from_string(
                        "DrawingML qualification mutation",
                    )))?;
                }
                edit.commit()
            })();
            NativeOwnerResult::DrawingmlChart(result)
        },
        (OwnerPrepared::DrawingmlThemeFamily(snapshot), OwnerRoute::DrawingmlThemeFamily) => {
            let result = (|| {
                let mut edit = snapshot.edit();
                if changed {
                    edit.set_name("DrawingML qualification mutation")?;
                }
                edit.commit()
            })();
            NativeOwnerResult::DrawingmlThemeFamily(result)
        },
        (OwnerPrepared::DrawingmlInkActions(profile), OwnerRoute::DrawingmlInkActions) => {
            use dml::ink::actions::{ActionSelector, Edit};
            let result = (|| {
                let mut edit = Edit::from_profile(profile)?;
                if changed {
                    edit.set_start_time(ActionSelector::ordinal(0), "0.50")?;
                }
                edit.commit()
            })();
            NativeOwnerResult::DrawingmlInkActions(result)
        },
        (OwnerPrepared::DrawingmlTransform(snapshot), OwnerRoute::DrawingmlTransform) => {
            let result = (|| {
                let mut edit = snapshot.edit();
                if changed {
                    edit.set_rotation(Some(dml::transform::Angle::new(60000)));
                }
                edit.commit()
            })();
            NativeOwnerResult::DrawingmlTransform(result)
        },
        _ => unreachable!("owner route and prepared source type differ"),
    }
}

/// Current owners expose immutable packages/prepared values.  In particular,
/// Ink uses `Prepared::edit()` so the timed operation does not rebuild a
/// detached profile from source bytes.
#[cfg(feature = "drawingml-owner-current")]
fn owner_native_commit_prepared(
    route: OwnerRoute,
    prepared: &OwnerPrepared,
    changed: bool,
) -> NativeOwnerResult {
    match (prepared, route) {
        (OwnerPrepared::DrawingmlChart(package), OwnerRoute::DrawingmlChart) => {
            let result = (|| {
                let mut edit = package.edit();
                if changed {
                    edit.set_title(Some(dml::chart::data::TitleText::from_string(
                        "DrawingML qualification mutation",
                    )))?;
                }
                edit.commit()
            })();
            NativeOwnerResult::DrawingmlChart(result)
        },
        (OwnerPrepared::DrawingmlThemeFamily(package), OwnerRoute::DrawingmlThemeFamily) => {
            let result = (|| {
                let mut edit = package.edit();
                if changed {
                    edit.set_name("DrawingML qualification mutation")?;
                }
                edit.commit()
            })();
            NativeOwnerResult::DrawingmlThemeFamily(result)
        },
        (OwnerPrepared::DrawingmlInkActions(prepared), OwnerRoute::DrawingmlInkActions) => {
            use dml::ink::actions::builder::ActionSelector;
            let result = (|| {
                let mut edit = prepared.edit()?;
                if changed {
                    edit.set_start_time(ActionSelector::ordinal(0), "0.50")?;
                }
                edit.commit()
            })();
            NativeOwnerResult::DrawingmlInkActions(result)
        },
        (OwnerPrepared::DrawingmlTransform(package), OwnerRoute::DrawingmlTransform) => {
            let result = (|| {
                let mut edit = package.edit();
                if changed {
                    edit.set_rotation(Some(dml::transform::Angle::new(60000)));
                }
                edit.commit()
            })();
            NativeOwnerResult::DrawingmlTransform(result)
        },
        _ => unreachable!("owner route and prepared source type differ"),
    }
}

fn owner_native_commit(work: &OwnerWork, changed: bool) -> NativeOwnerResult {
    owner_native_commit_prepared(work.fixture.route, &work.prepared, changed)
}

#[cfg(not(feature = "drawingml-owner-current"))]
fn owner_native_replay(prepared: &OwnerPrepared, patch: &OwnerPatch) -> NativeReplayResult {
    match (prepared, patch) {
        (OwnerPrepared::DrawingmlChart(snapshot), OwnerPatch::DrawingmlChart(patch)) => {
            NativeReplayResult::DrawingmlChart(patch.apply(snapshot))
        },
        (
            OwnerPrepared::DrawingmlThemeFamily(snapshot),
            OwnerPatch::DrawingmlThemeFamily(patch),
        ) => NativeReplayResult::DrawingmlThemeFamily(patch.apply(snapshot)),
        (OwnerPrepared::DrawingmlInkActions(profile), OwnerPatch::DrawingmlInkActions(patch)) => {
            NativeReplayResult::DrawingmlInkActions(patch.apply(profile.source()))
        },
        (OwnerPrepared::DrawingmlTransform(snapshot), OwnerPatch::DrawingmlTransform(patch)) => {
            NativeReplayResult::DrawingmlTransform(patch.apply(snapshot))
        },
        _ => unreachable!("owner route, prepared source and patch type differ"),
    }
}

#[cfg(feature = "drawingml-owner-current")]
fn owner_native_replay(prepared: &OwnerPrepared, patch: &OwnerPatch) -> NativeReplayResult {
    match (prepared, patch) {
        (OwnerPrepared::DrawingmlChart(package), OwnerPatch::DrawingmlChart(patch)) => {
            NativeReplayResult::DrawingmlChart(package.apply(patch))
        },
        (OwnerPrepared::DrawingmlThemeFamily(package), OwnerPatch::DrawingmlThemeFamily(patch)) => {
            NativeReplayResult::DrawingmlThemeFamily(package.apply(patch))
        },
        (OwnerPrepared::DrawingmlInkActions(prepared), OwnerPatch::DrawingmlInkActions(patch)) => {
            NativeReplayResult::DrawingmlInkActions(prepared.apply(patch))
        },
        (OwnerPrepared::DrawingmlTransform(package), OwnerPatch::DrawingmlTransform(patch)) => {
            NativeReplayResult::DrawingmlTransform(package.apply(patch))
        },
        _ => unreachable!("owner route, prepared source and patch type differ"),
    }
}

fn black_box_native_owner(result: NativeOwnerResult) {
    match result {
        NativeOwnerResult::DrawingmlChart(result) => {
            let _ = black_box(result);
        },
        NativeOwnerResult::DrawingmlThemeFamily(result) => {
            let _ = black_box(result);
        },
        NativeOwnerResult::DrawingmlInkActions(result) => {
            let _ = black_box(result);
        },
        NativeOwnerResult::DrawingmlTransform(result) => {
            let _ = black_box(result);
        },
    }
}

fn black_box_native_replay(result: NativeReplayResult) {
    match result {
        NativeReplayResult::DrawingmlChart(result) => {
            let _ = black_box(result);
        },
        NativeReplayResult::DrawingmlThemeFamily(result) => {
            let _ = black_box(result);
        },
        NativeReplayResult::DrawingmlInkActions(result) => {
            let _ = black_box(result);
        },
        NativeReplayResult::DrawingmlTransform(result) => {
            let _ = black_box(result);
        },
    }
}

#[cfg(not(feature = "drawingml-owner-current"))]
fn owner_native_commit_parts(
    route: OwnerRoute,
    result: NativeOwnerResult,
) -> Result<(Vec<u8>, OwnerPatch), ProbeFailure> {
    match (route, result) {
        (OwnerRoute::DrawingmlChart, NativeOwnerResult::DrawingmlChart(result)) => result
            .map(|commit| {
                (
                    commit.snapshot().xml_bytes().to_vec(),
                    OwnerPatch::DrawingmlChart(commit.patch().clone()),
                )
            })
            .map_err(|error| owner_error(route, error)),
        (OwnerRoute::DrawingmlThemeFamily, NativeOwnerResult::DrawingmlThemeFamily(result)) => {
            result
                .map(|commit| {
                    (
                        commit.snapshot().xml_bytes().to_vec(),
                        OwnerPatch::DrawingmlThemeFamily(commit.patch().clone()),
                    )
                })
                .map_err(|error| owner_error(route, error))
        },
        (OwnerRoute::DrawingmlInkActions, NativeOwnerResult::DrawingmlInkActions(result)) => result
            .map(|commit| {
                (
                    commit.as_bytes().to_vec(),
                    OwnerPatch::DrawingmlInkActions(commit.patch().clone()),
                )
            })
            .map_err(|error| owner_error(route, error)),
        (OwnerRoute::DrawingmlTransform, NativeOwnerResult::DrawingmlTransform(result)) => result
            .map(|commit| {
                (
                    commit.snapshot().xml_bytes().to_vec(),
                    OwnerPatch::DrawingmlTransform(commit.patch().clone()),
                )
            })
            .map_err(|error| owner_error(route, error)),
        _ => Err(fail("owner route and native commit result type differ")),
    }
}

#[cfg(feature = "drawingml-owner-current")]
fn owner_native_commit_parts(
    route: OwnerRoute,
    result: NativeOwnerResult,
) -> Result<(Vec<u8>, OwnerPatch), ProbeFailure> {
    match (route, result) {
        (OwnerRoute::DrawingmlChart, NativeOwnerResult::DrawingmlChart(result)) => result
            .map(|commit| {
                (
                    commit.package().xml_bytes().to_vec(),
                    OwnerPatch::DrawingmlChart(commit.patch().clone()),
                )
            })
            .map_err(|error| owner_error(route, error)),
        (OwnerRoute::DrawingmlThemeFamily, NativeOwnerResult::DrawingmlThemeFamily(result)) => {
            result
                .map(|commit| {
                    (
                        commit.package().xml_bytes().to_vec(),
                        OwnerPatch::DrawingmlThemeFamily(commit.patch().clone()),
                    )
                })
                .map_err(|error| owner_error(route, error))
        },
        (OwnerRoute::DrawingmlInkActions, NativeOwnerResult::DrawingmlInkActions(result)) => result
            .map(|commit| {
                (
                    commit.package().as_bytes().to_vec(),
                    OwnerPatch::DrawingmlInkActions(commit.patch().clone()),
                )
            })
            .map_err(|error| owner_error(route, error)),
        (OwnerRoute::DrawingmlTransform, NativeOwnerResult::DrawingmlTransform(result)) => result
            .map(|commit| {
                (
                    commit.package().xml_bytes().to_vec(),
                    OwnerPatch::DrawingmlTransform(commit.patch().clone()),
                )
            })
            .map_err(|error| owner_error(route, error)),
        _ => Err(fail("owner route and native commit result type differ")),
    }
}

fn owner_native_commit_output(
    route: OwnerRoute,
    result: NativeOwnerResult,
) -> Result<Vec<u8>, ProbeFailure> {
    owner_native_commit_parts(route, result).map(|(output, _patch)| output)
}

#[cfg(not(feature = "drawingml-owner-current"))]
fn owner_native_replay_output(
    route: OwnerRoute,
    result: NativeReplayResult,
) -> Result<Vec<u8>, ProbeFailure> {
    match (route, result) {
        (OwnerRoute::DrawingmlChart, NativeReplayResult::DrawingmlChart(result)) => result
            .map(|snapshot| snapshot.xml_bytes().to_vec())
            .map_err(|error| owner_error(route, error)),
        (OwnerRoute::DrawingmlThemeFamily, NativeReplayResult::DrawingmlThemeFamily(result)) => {
            result
                .map(|snapshot| snapshot.xml_bytes().to_vec())
                .map_err(|error| owner_error(route, error))
        },
        (OwnerRoute::DrawingmlInkActions, NativeReplayResult::DrawingmlInkActions(result)) => {
            result.map_err(|error| owner_error(route, error))
        },
        (OwnerRoute::DrawingmlTransform, NativeReplayResult::DrawingmlTransform(result)) => result
            .map(|snapshot| snapshot.xml_bytes().to_vec())
            .map_err(|error| owner_error(route, error)),
        _ => Err(fail("owner route and native replay result type differ")),
    }
}

#[cfg(feature = "drawingml-owner-current")]
fn owner_native_replay_output(
    route: OwnerRoute,
    result: NativeReplayResult,
) -> Result<Vec<u8>, ProbeFailure> {
    match (route, result) {
        (OwnerRoute::DrawingmlChart, NativeReplayResult::DrawingmlChart(result)) => result
            .map(|commit| commit.package().xml_bytes().to_vec())
            .map_err(|error| owner_error(route, error)),
        (OwnerRoute::DrawingmlThemeFamily, NativeReplayResult::DrawingmlThemeFamily(result)) => {
            result
                .map(|commit| commit.package().xml_bytes().to_vec())
                .map_err(|error| owner_error(route, error))
        },
        (OwnerRoute::DrawingmlInkActions, NativeReplayResult::DrawingmlInkActions(result)) => {
            result
                .map(|commit| commit.package().as_bytes().to_vec())
                .map_err(|error| owner_error(route, error))
        },
        (OwnerRoute::DrawingmlTransform, NativeReplayResult::DrawingmlTransform(result)) => result
            .map(|commit| commit.package().xml_bytes().to_vec())
            .map_err(|error| owner_error(route, error)),
        _ => Err(fail("owner route and native replay result type differ")),
    }
}

fn owner_native_replay_error(
    route: OwnerRoute,
    result: NativeReplayResult,
) -> Result<ProbeFailure, ProbeFailure> {
    match result {
        NativeReplayResult::DrawingmlChart(Ok(_))
        | NativeReplayResult::DrawingmlThemeFamily(Ok(_))
        | NativeReplayResult::DrawingmlInkActions(Ok(_))
        | NativeReplayResult::DrawingmlTransform(Ok(_)) => Err(fail(format!(
            "{} accepted a stale patch",
            owner_route_name(route)
        ))),
        NativeReplayResult::DrawingmlChart(Err(error))
        | NativeReplayResult::DrawingmlThemeFamily(Err(error))
        | NativeReplayResult::DrawingmlInkActions(Err(error))
        | NativeReplayResult::DrawingmlTransform(Err(error)) => Ok(owner_error(route, error)),
    }
}

fn validate_stale_refusal(
    route: OwnerRoute,
    error: ProbeFailure,
    observed: usize,
) -> Result<RefusalReceipt, ProbeFailure> {
    let Some(mut refusal) = error.refusal else {
        return Err(fail(format!(
            "{} stale replay returned an untyped error: {}",
            owner_route_name(route),
            error.message
        )));
    };
    #[cfg(feature = "drawingml-owner-current")]
    if refusal.kind != "patch_conflict" || refusal.source_kind != Some("patch_conflict") {
        return Err(fail(format!(
            "{} stale replay returned the wrong typed error: {}",
            owner_route_name(route),
            refusal.message
        )));
    }
    #[cfg(not(feature = "drawingml-owner-current"))]
    if refusal.kind != "patch_conflict" || refusal.source_kind != Some("invalid_input") {
        return Err(fail(format!(
            "{} stale replay returned the wrong legacy classification: {}",
            owner_route_name(route),
            refusal.message
        )));
    }
    refusal.observed = observed;
    Ok(refusal)
}

fn owner_native_stale_refusal(work: &OwnerWork) -> Result<RefusalReceipt, ProbeFailure> {
    let error = match owner_native_replay_error(
        work.fixture.route,
        owner_native_replay(&work.stale_prepared, &work.patch),
    ) {
        Ok(error) => error,
        Err(error) => return Err(error),
    };
    validate_stale_refusal(work.fixture.route, error, work.fixture.stale.len())
}

#[cfg(not(feature = "drawingml-owner-current"))]
const fn stale_error_variant() -> &'static str {
    "litchi_drawingml::Error::Invalid"
}

#[cfg(feature = "drawingml-owner-current")]
const fn stale_error_variant() -> &'static str {
    "litchi_drawingml::Error::PatchConflict"
}

#[cfg(not(feature = "drawingml-owner-current"))]
fn stale_reason_verified(route: OwnerRoute, refusal: &RefusalReceipt) -> bool {
    let reason = match route {
        OwnerRoute::DrawingmlChart => "chart patch source does not match its byte precondition",
        OwnerRoute::DrawingmlThemeFamily => {
            "theme family patch source does not match its byte precondition"
        },
        OwnerRoute::DrawingmlInkActions => "ink action patch source is stale",
        OwnerRoute::DrawingmlTransform => {
            "DrawingML transform patch source state does not match its precondition"
        },
    };
    // Legacy `Error::Invalid` retains the exact source-specific cause inside
    // the public Display prefix.  Verify both the concrete variant (recorded
    // separately) and that full, non-redacted reason here.
    refusal.message == format!("invalid DrawingML: {reason}")
}

#[cfg(feature = "drawingml-owner-current")]
fn stale_reason_verified(_route: OwnerRoute, refusal: &RefusalReceipt) -> bool {
    refusal.message == "DrawingML patch source does not match the recipient"
}

fn owner_proof(
    fixture_identity_sha256: &str,
    work: &OwnerWork,
) -> Result<OwnerCommitReceipt, ProbeFailure> {
    let fixture = &work.fixture;
    // Correctness uses the very same native helpers as the timed rows.  Only
    // the native result-to-bytes conversion and semantic oracle run here,
    // outside the measurement clock.
    let noop_bytes = owner_native_commit_output(fixture.route, owner_native_commit(work, false))?;
    if noop_bytes != fixture.source {
        return Err(fail(format!(
            "{} no-op commit changed the exact source bytes",
            owner_route_name(fixture.route)
        )));
    }
    let changed_bytes = owner_native_commit_output(fixture.route, owner_native_commit(work, true))?;
    if changed_bytes != work.changed {
        return Err(fail(format!(
            "{} native changed commit differs from its prepared owner output",
            owner_route_name(fixture.route)
        )));
    }
    let forward_bytes = owner_native_replay_output(
        fixture.route,
        owner_native_replay(&work.prepared, &work.patch),
    )?;
    let inverse_bytes = owner_native_replay_output(
        fixture.route,
        owner_native_replay(&work.changed_prepared, &work.inverse_patch),
    )?;
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
    let stale_refusal = owner_native_stale_refusal(work)?;
    if !stale_reason_verified(fixture.route, &stale_refusal) {
        return Err(fail(format!(
            "{} stale refusal did not preserve its exact source reason",
            owner_route_name(fixture.route)
        )));
    }
    let semantic = owner_semantic_receipt(
        fixture.route,
        &fixture.source,
        &noop_bytes,
        &changed_bytes,
        &forward_bytes,
        &inverse_bytes,
    )?;
    Ok(OwnerCommitReceipt {
        route: owner_route_name(fixture.route),
        native_commit_path: owner_operation_path(fixture.route, "changed_commit"),
        native_replay_path: owner_operation_path(fixture.route, "forward_replay"),
        fixture_identity_sha256: sha256(
            format!(
                "drawingml-owner-fixture-v1\0{fixture_identity_sha256}\0{}\0{}\0{}",
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
        stale_error_variant: stale_error_variant(),
        stale_reason_verified: true,
        stale_refusal,
        semantic,
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
            format!(
                "text={:?};paragraphs={}",
                body.text(),
                body.paragraphs.len()
            )
        },
        PartKind::SvgBlip => {
            let value = read_svg_value(bytes)?;
            format!(
                // Contextual export may materialize inherited declarations on
                // the standalone root.  Compare the effective typed/opaque
                // payload, not the spelling of that namespace completion.
                "reference={:?};attributes={:?};children={:?}",
                value.reference(),
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
            format!(
                "has_family={};uri={:?};family={family:?}",
                value.has_family(),
                value.family_extension_uri()
            )
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

fn owner_part_kind(route: OwnerRoute) -> PartKind {
    match route {
        OwnerRoute::DrawingmlChart => PartKind::Chart,
        OwnerRoute::DrawingmlThemeFamily => PartKind::ThemeFamily,
        OwnerRoute::DrawingmlInkActions => PartKind::InkActions,
        OwnerRoute::DrawingmlTransform => PartKind::Transform,
    }
}

fn first_ink_action<'a>(
    children: &'a [dml::ink::actions::RootChild],
) -> Option<&'a dml::ink::actions::Action> {
    for child in children {
        match child {
            dml::ink::actions::RootChild::Action(action) => return Some(action),
            dml::ink::actions::RootChild::ActionGroup(group) => {
                if let Some(action) = group.actions().first() {
                    return Some(action);
                }
            },
            _ => {},
        }
    }
    None
}

/// Return a semantic remainder fingerprint after removing the field changed
/// by the owner fixture.  The full semantic fingerprint remains separate and
/// proves the mutation itself; this fingerprint proves that the requested
/// setter did not silently discard the other modeled fields or retained raw
/// profile content.
fn owner_rest_semantic_key(route: OwnerRoute, bytes: &[u8]) -> Result<String, ProbeFailure> {
    match route {
        OwnerRoute::DrawingmlChart => {
            let mut value = read_chart_value(bytes)?;
            value.title = None;
            Ok(sha256(format!("{:?}", value).as_bytes()))
        },
        OwnerRoute::DrawingmlThemeFamily => {
            let value = read_family_value(bytes)?;
            Ok(sha256(
                format!("{:?}|{:?}", value.id(), value.variant_id()).as_bytes(),
            ))
        },
        OwnerRoute::DrawingmlInkActions => {
            let profile = read_ink_value(bytes)?;
            let source = profile.source();
            let Some(action) = first_ink_action(profile.children()) else {
                return Err(fail("Ink owner fixture has no first action"));
            };
            let (start, end) = action.source_span();
            if start > end || end > source.len() {
                return Err(fail("Ink owner first action span is outside its source"));
            }
            let mut normalized = Vec::with_capacity(source.len() + 16);
            normalized.extend_from_slice(&source[..start]);
            normalized.extend(normalize_ink_start_time(&source[start..end]));
            normalized.extend_from_slice(&source[end..]);
            Ok(sha256(&normalized))
        },
        OwnerRoute::DrawingmlTransform => {
            let mut value = read_transform_value(bytes)?;
            value.set_rotation(None);
            Ok(sha256(format!("{:?}", value).as_bytes()))
        },
    }
}

fn normalize_ink_start_time(action: &[u8]) -> Vec<u8> {
    let needle = b"startTime";
    let Some(name_start) = action
        .windows(needle.len())
        .position(|window| window == needle)
    else {
        return action.to_vec();
    };
    let Some(equal_offset) = action[name_start + needle.len()..]
        .iter()
        .position(|byte| *byte == b'=')
    else {
        return action.to_vec();
    };
    let equal = name_start + needle.len() + equal_offset;
    let Some(quote) = action.get(equal + 1).copied() else {
        return action.to_vec();
    };
    if quote != b'"' && quote != b'\'' {
        return action.to_vec();
    }
    let value_start = equal + 2;
    let Some(value_len) = action[value_start..].iter().position(|byte| *byte == quote) else {
        return action.to_vec();
    };
    let value_end = value_start + value_len;
    let mut normalized = Vec::with_capacity(action.len() + 16);
    normalized.extend_from_slice(&action[..value_start]);
    normalized.extend_from_slice(b"<owner-start-time>");
    normalized.extend_from_slice(&action[value_end..]);
    normalized
}

fn owner_semantic_receipt(
    route: OwnerRoute,
    source: &[u8],
    noop: &[u8],
    changed: &[u8],
    forward: &[u8],
    inverse: &[u8],
) -> Result<OwnerSemanticReceipt, ProbeFailure> {
    let kind = owner_part_kind(route);
    let source_key = semantic_key(kind, source)?;
    let noop_key = semantic_key(kind, noop)?;
    let changed_key = semantic_key(kind, changed)?;
    let forward_key = semantic_key(kind, forward)?;
    let inverse_key = semantic_key(kind, inverse)?;
    if source_key != noop_key {
        return Err(fail(format!(
            "{} no-op semantic projection changed",
            owner_route_name(route)
        )));
    }
    if source_key == changed_key || forward_key != changed_key || inverse_key != source_key {
        return Err(fail(format!(
            "{} semantic commit/replay identities are not reversible",
            owner_route_name(route)
        )));
    }
    let (expected_mutation, mutation_verified) = match route {
        OwnerRoute::DrawingmlChart => {
            let value = read_chart_value(changed)?;
            let verified = matches!(
                value.title.as_ref(),
                Some(dml::chart::data::TitleText::Literal(text))
                    if text.text == "DrawingML qualification mutation"
            );
            ("chart.title=DrawingML qualification mutation", verified)
        },
        OwnerRoute::DrawingmlThemeFamily => {
            let value = read_family_value(changed)?;
            (
                "theme_family.name=DrawingML qualification mutation",
                value.name() == "DrawingML qualification mutation",
            )
        },
        OwnerRoute::DrawingmlInkActions => {
            let profile = read_ink_value(changed)?;
            let verified = first_ink_action(profile.children())
                .is_some_and(|action| action.start_time() == "0.50");
            ("ink_actions.action[0].start_time=0.50", verified)
        },
        OwnerRoute::DrawingmlTransform => {
            let value = read_transform_value(changed)?;
            (
                "transform.rotation=60000",
                value.authored_rotation() == Some(dml::transform::Angle::new(60000)),
            )
        },
    };
    if !mutation_verified {
        return Err(fail(format!(
            "{} changed semantic projection does not contain the expected mutation",
            owner_route_name(route)
        )));
    }
    let source_rest = owner_rest_semantic_key(route, source)?;
    let changed_rest = owner_rest_semantic_key(route, changed)?;
    let unchanged_remainder = source_rest == changed_rest;
    if !unchanged_remainder {
        return Err(fail(format!(
            "{} mutation changed an unrelated semantic remainder",
            owner_route_name(route)
        )));
    }
    Ok(OwnerSemanticReceipt {
        source_semantic_sha256: sha256(source_key.as_bytes()),
        noop_semantic_sha256: sha256(noop_key.as_bytes()),
        changed_semantic_sha256: sha256(changed_key.as_bytes()),
        forward_semantic_sha256: sha256(forward_key.as_bytes()),
        inverse_semantic_sha256: sha256(inverse_key.as_bytes()),
        source_rest_semantic_sha256: source_rest,
        changed_rest_semantic_sha256: changed_rest,
        expected_mutation,
        mutation_verified,
        unchanged_remainder,
    })
}

fn semantic_equal(kind: PartKind, left: &[u8], right: &[u8]) -> Result<bool, ProbeFailure> {
    Ok(match kind {
        PartKind::Chart => {
            format!("{:?}", read_chart_value(left)?) == format!("{:?}", read_chart_value(right)?)
        },
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

/// Prove a changed source-backed fixture retained its untouched lexical
/// slices without asking the production writer to generate an expected byte
/// string.  The replacement region is deliberately treated as opaque here;
/// semantic readback and the requested mutation prove what changed in that
/// region separately.
fn splice_oracle(source: &[u8], changed: &[u8], marker: &[u8]) -> bool {
    fn root_span(bytes: &[u8], marker: &[u8]) -> Option<(usize, usize)> {
        let start = bytes
            .windows(marker.len())
            .position(|window| window == marker)?;
        let tail = &bytes[start..];
        let end = if let Some(relative) = tail.windows(2).position(|window| window == b"/>") {
            start + relative + 2
        } else {
            let close = b"</thm15:themeFamily>";
            let relative = tail
                .windows(close.len())
                .position(|window| window == close)?;
            start + relative + close.len()
        };
        Some((start, end))
    }

    let (source_start, source_end) = if marker == b"<![CDATA[" {
        let Some(start) = source
            .windows(marker.len())
            .position(|window| window == marker)
        else {
            return false;
        };
        let value_start = start + marker.len();
        let suffix = b"]]>";
        let Some(relative) = source[value_start..]
            .windows(suffix.len())
            .position(|window| window == suffix)
        else {
            return false;
        };
        (value_start, value_start + relative)
    } else {
        let Some(span) = root_span(source, marker) else {
            return false;
        };
        span
    };
    let (changed_start, changed_end) = if marker == b"<![CDATA[" {
        let Some(start) = changed
            .windows(marker.len())
            .position(|window| window == marker)
        else {
            return false;
        };
        let value_start = start + marker.len();
        let suffix = b"]]>";
        let Some(relative) = changed[value_start..]
            .windows(suffix.len())
            .position(|window| window == suffix)
        else {
            return false;
        };
        (value_start, value_start + relative)
    } else {
        let Some(span) = root_span(changed, marker) else {
            return false;
        };
        span
    };
    source[..source_start] == changed[..changed_start]
        && source[source_end..] == changed[changed_end..]
        && source[source_start..source_end] != changed[changed_start..changed_end]
}

fn independent_splice_oracle(kind: PartKind, source: &[u8], changed: &[u8]) -> (bool, bool) {
    match kind {
        PartKind::FormatCode => (true, splice_oracle(source, changed, b"<![CDATA[")),
        PartKind::ThemeFamily | PartKind::ThemeFamilyPart => {
            (true, splice_oracle(source, changed, b"<thm15:themeFamily"))
        },
        _ => (false, true),
    }
}

fn raw_fidelity(case: &LaneCase) -> Result<Vec<RawFidelityReceipt>, ProbeFailure> {
    case.parts
        .iter()
        .map(|part| {
            let (source_output, changed_output, read_write_route) =
                if part.kind == PartKind::SvgBlip {
                    let prepared = case
                        .prepared
                        .iter()
                        .find(|value| value.id == part.id)
                        .ok_or_else(|| fail("SVG contextual fixture was not prepared"))?;
                    let source_value = prepared
                        .contextual_source
                        .as_ref()
                        .ok_or_else(|| fail("SVG contextual source was not prepared"))?;
                    let changed_value = prepared
                        .contextual_changed
                        .as_ref()
                        .ok_or_else(|| fail("SVG contextual changed value was not prepared"))?;
                    (
                        write_svg_contextual_default(source_value)?,
                        write_svg_contextual_default(changed_value)?,
                        "contextual_read_write",
                    )
                } else {
                    (
                        roundtrip(part.kind, &part.source)?,
                        roundtrip(part.kind, &part.changed)?,
                        "standalone_read_write",
                    )
                };
            if source_output.is_empty() || changed_output.is_empty() {
                return Err(fail(format!("{} raw readback was empty", part.id)));
            }
            if part.kind == PartKind::SvgBlip {
                let prepared = case
                    .prepared
                    .iter()
                    .find(|value| value.id == part.id)
                    .ok_or_else(|| fail("SVG contextual fixture was not prepared"))?;
                if prepared.contextual_source_output_bytes != Some(source_output.len())
                    || prepared.contextual_changed_output_bytes != Some(changed_output.len())
                {
                    return Err(fail(
                        "SVG contextual output changed after the pre-warmup handoff",
                    ));
                }
            }
            let standalone_route_checked = if part.kind == PartKind::SvgBlip {
                let standalone_source = standalone_svg_roundtrip(&source_output)?;
                let standalone_changed = standalone_svg_roundtrip(&changed_output)?;
                if !semantic_equal(PartKind::SvgBlip, &source_output, &standalone_source)?
                    || !semantic_equal(PartKind::SvgBlip, &changed_output, &standalone_changed)?
                {
                    return Err(fail(
                        "standalone SVG route changed the contextual semantic model",
                    ));
                }
                true
            } else {
                false
            };
            let source_semantic = semantic_key(part.kind, &part.source)?;
            let source_output_semantic = semantic_key(part.kind, &source_output)?;
            let changed_semantic = semantic_key(part.kind, &part.changed)?;
            let changed_output_semantic = semantic_key(part.kind, &changed_output)?;
            let source_semantic_equal = semantic_equal(part.kind, &part.source, &source_output)?;
            let changed_semantic_equal = semantic_equal(part.kind, &part.changed, &changed_output)?;
            let semantic_changed_differs = !semantic_equal(part.kind, &part.source, &part.changed)?;
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
            let source_sha256 = sha256(&part.source);
            let changed_sha256 = sha256(&part.changed);
            let malformed_sha256 = sha256(&part.malformed);
            let source_preserving_expected = source_preserving(part.kind);
            let source_raw_sha256 = if part.kind == PartKind::SvgBlip {
                let prepared = case
                    .prepared
                    .iter()
                    .find(|value| value.id == part.id)
                    .ok_or_else(|| fail("SVG contextual fixture was not prepared"))?;
                let source_value = prepared
                    .contextual_source
                    .as_ref()
                    .ok_or_else(|| fail("SVG contextual source was not prepared"))?;
                let raw = source_value
                    .raw_source()
                    .ok_or_else(|| fail("SVG contextual source lost its raw fragment"))?;
                sha256(raw)
            } else {
                source_sha256.clone()
            };
            let source_raw_exact = if part.kind == PartKind::SvgBlip {
                let prepared = case
                    .prepared
                    .iter()
                    .find(|value| value.id == part.id)
                    .ok_or_else(|| fail("SVG contextual fixture was not prepared"))?;
                let source_value = prepared
                    .contextual_source
                    .as_ref()
                    .ok_or_else(|| fail("SVG contextual source was not prepared"))?;
                source_value.raw_source() == Some(part.source.as_slice())
            } else {
                source_output == part.source
            };
            if source_preserving_expected && !source_raw_exact {
                return Err(fail(format!(
                    "{} canonical source changed under a promised source-preserving route",
                    part.id
                )));
            }
            let changed_output_exact = changed_output == part.changed;
            let changed_output_exact_expected =
                source_preserving_expected && part.kind != PartKind::SvgBlip;
            if changed_output_exact_expected && !changed_output_exact {
                return Err(fail(format!(
                    "{} changed source changed under a promised source-preserving route",
                    part.id
                )));
            }
            let (splice_oracle_checked, splice_oracle_passed) =
                independent_splice_oracle(part.kind, &part.source, &part.changed);
            if splice_oracle_checked && !splice_oracle_passed {
                return Err(fail(format!(
                    "{} changed fixture did not preserve its independent untouched slices",
                    part.id
                )));
            }
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
                changed_output_exact,
                source_raw_sha256,
                source_raw_exact,
                changed_output_exact_expected,
                standalone_route_checked,
                splice_oracle_checked,
                splice_oracle_passed,
                read_write_route,
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
        PartKind::ThemeFamily => Some(("theme family XML bytes", 1 << 20)),
        PartKind::ThemeFamilyPart => Some((
            "complete Theme-part XML bytes",
            dml::theme::family::part::MAX_XML_BYTES,
        )),
        PartKind::InkActions => Some(("ink actions source bytes", 16 * 1024 * 1024)),
        PartKind::Color => Some(("DrawingML color XML", 64 * 1024)),
        PartKind::Transform => Some(("DrawingML transform XML", 1 << 20)),
        PartKind::Blip | PartKind::Geometry => None,
    }
}

fn over_limit(kind: PartKind, source: &[u8]) -> Vec<u8> {
    let (_, limit) = input_limit(kind).expect("fixture boundary requires an input limit");
    // Keep the boundary fixture valid so a typed input-limit refusal is not
    // hidden behind an earlier malformed-XML refusal.  XML permits trailing
    // whitespace after the complete document.
    sized_with_whitespace(source, limit.saturating_add(1))
}

fn part_has_input_limit(kind: PartKind) -> bool {
    input_limit(kind).is_some()
}

fn geometry_invalid(message: impl Into<String>) -> dml::Error {
    dml::Error::Invalid(message.into())
}

/// Parse custom geometry with the native DrawingML error boundary intact.
/// The timed entry row calls this function directly and black-boxes its
/// `dml::Result`; only the untimed detail pass maps that result into the
/// probe's refusal receipt.
fn parse_geometry_native(bytes: &[u8]) -> dml::Result<dml::geometry::CustomGeometry> {
    use dml::geometry::reader::{CustomGeometryBuilder, GeometryElement};
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    #[cfg(feature = "drawingml-raw-current")]
    let mut builder = CustomGeometryBuilder::new(dml::geometry::Limits::DEFAULT)?;
    #[cfg(not(feature = "drawingml-raw-current"))]
    let mut builder = CustomGeometryBuilder::new();
    let mut stack = Vec::new();
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(dml::Error::from)?
            .into_owned();
        match event {
            Event::Start(element) => {
                let local = element.local_name().as_ref().to_vec();
                if stack.is_empty() {
                    if local.as_slice() != b"custGeom" {
                        return Err(geometry_invalid("geometry root is not custGeom"));
                    }
                    stack.push(GeometryElement::CustomGeometry);
                } else {
                    let parent = *stack
                        .last()
                        .ok_or_else(|| geometry_invalid("geometry stack is empty"))?;
                    let child = builder.open(parent, &local, &element, reader.decoder())?;
                    let child = child.ok_or_else(|| {
                        geometry_invalid(format!(
                            "unsupported geometry child {:?}",
                            String::from_utf8_lossy(&local)
                        ))
                    })?;
                    stack.push(child);
                }
            },
            Event::Empty(element) => {
                let local = element.local_name().as_ref().to_vec();
                if stack.is_empty() {
                    return Err(geometry_invalid("empty geometry root is unsupported"));
                }
                let parent = *stack
                    .last()
                    .ok_or_else(|| geometry_invalid("geometry stack is empty"))?;
                let child = builder.open(parent, &local, &element, reader.decoder())?;
                let child = child.ok_or_else(|| {
                    geometry_invalid(format!(
                        "unsupported geometry child {:?}",
                        String::from_utf8_lossy(&local)
                    ))
                })?;
                builder.close(child)?;
            },
            Event::End(_) => {
                let element = stack
                    .pop()
                    .ok_or_else(|| geometry_invalid("geometry close underflow"))?;
                builder.close(element)?;
                if stack.is_empty() {
                    break;
                }
            },
            Event::Eof => return Err(geometry_invalid("geometry root is unterminated")),
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_) => {},
            Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(geometry_invalid("geometry contains forbidden markup"));
            },
        }
        buffer.clear();
    }
    builder.finish()
}

fn parse_geometry(bytes: &[u8]) -> Result<dml::geometry::CustomGeometry, ProbeFailure> {
    parse_geometry_native(bytes).map_err(|error| classify(error, bytes.len()))
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
            let value = dml::chart::extension::formatcode2::read_shared(std::sync::Arc::clone(
                &part.source,
            ));
            Ok(PreparedValue::FormatCode(
                value.map_err(|error| classify(error, part.source.len()))?,
            ))
        },
        PartKind::ThemeFamily => {
            #[cfg(feature = "drawingml-raw-current")]
            let value = dml::theme::family::read_shared(
                std::sync::Arc::clone(&part.source),
                dml::theme::family::Limits::DEFAULT,
            );
            #[cfg(not(feature = "drawingml-raw-current"))]
            let value = dml::theme::family::read_shared(std::sync::Arc::clone(&part.source));
            Ok(PreparedValue::ThemeFamily(
                value.map_err(|error| classify(error, part.source.len()))?,
            ))
        },
        PartKind::ThemeFamilyPart => {
            #[cfg(feature = "drawingml-raw-current")]
            let value = dml::theme::family::part::read_shared(
                std::sync::Arc::clone(&part.source),
                dml::theme::family::part::Limits::DEFAULT,
            );
            #[cfg(not(feature = "drawingml-raw-current"))]
            let value = dml::theme::family::part::read_shared(std::sync::Arc::clone(&part.source));
            let snapshot = value.map_err(|error| classify(error, part.source.len()))?;
            let family = snapshot
                .family()
                .cloned()
                .ok_or_else(|| fail("shared Theme part has no direct family owner"))?;
            Ok(PreparedValue::ThemeFamilyPart { snapshot, family })
        },
        _ => Err(fail(format!("{} has no shared-reader route", part.id))),
    }
}

fn timed_shared_read(case: &LaneCase) {
    for part in &case.shared {
        match part.kind {
            PartKind::FormatCode => {
                #[cfg(feature = "drawingml-raw-current")]
                black_box(dml::chart::extension::formatcode2::read_shared(
                    std::sync::Arc::clone(&part.source),
                    &dml::chart::extension::formatcode2::Limits::DEFAULT,
                ));
                #[cfg(not(feature = "drawingml-raw-current"))]
                black_box(dml::chart::extension::formatcode2::read_shared(
                    std::sync::Arc::clone(&part.source),
                ));
            },
            PartKind::ThemeFamily => {
                #[cfg(feature = "drawingml-raw-current")]
                black_box(dml::theme::family::read_shared(
                    std::sync::Arc::clone(&part.source),
                    dml::theme::family::Limits::DEFAULT,
                ));
                #[cfg(not(feature = "drawingml-raw-current"))]
                black_box(dml::theme::family::read_shared(std::sync::Arc::clone(
                    &part.source,
                )));
            },
            PartKind::ThemeFamilyPart => {
                #[cfg(feature = "drawingml-raw-current")]
                black_box(dml::theme::family::part::read_shared(
                    std::sync::Arc::clone(&part.source),
                    dml::theme::family::part::Limits::DEFAULT,
                ));
                #[cfg(not(feature = "drawingml-raw-current"))]
                black_box(dml::theme::family::part::read_shared(
                    std::sync::Arc::clone(&part.source),
                ));
            },
            _ => {},
        }
    }
}

fn shared_operation(case: &LaneCase, timed: bool) -> Result<OperationValue, ProbeFailure> {
    if timed {
        // Keep the production shared reader result opaque to the timer.  The
        // parsed value and receipt validation belong to the pre-warmup detail
        // pass; this row measures only the shared reader call itself.
        timed_shared_read(case);
        return Ok(OperationValue::Timed);
    }
    let mut sink = CountingSink::new(usize::MAX);
    for value in &case.shared_prepared {
        let output = match &value {
            PreparedValue::FormatCode(value) => {
                let mut output = Vec::new();
                write_format_code_value(&mut output, value)?;
                output
            },
            PreparedValue::ThemeFamily(value) => write_family_value(value)?,
            PreparedValue::ThemeFamilyPart { snapshot, .. } => snapshot.xml_bytes().to_vec(),
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
    let source_family = source_snapshot
        .family()
        .cloned()
        .ok_or_else(|| fail("source Theme part has no direct family owner"))?;
    let (expected_source_name, expected_source_id, expected_source_variant_id) =
        theme_family_fixture_identity();
    let source_identity_matches = source_family.name() == expected_source_name
        && source_family.id().to_string() == expected_source_id
        && source_family.variant_id().to_string() == expected_source_variant_id;
    if !source_identity_matches {
        return Err(fail(
            "Theme-part mutation source identity does not match its bound name/id/variant",
        ));
    }
    let source_name = Some(source_family.name().to_owned());
    let source_id = Some(source_family.id().to_string());
    let source_variant_id = Some(source_family.variant_id().to_string());
    let source_sha256 = sha256(&part.source);
    let replaced = theme_family_replace(&source_snapshot, &changed_family)
        .map_err(|error| classify(error, part.source.len()))?;
    if replaced == part.source {
        return Err(fail(
            "theme-family replace did not change the complete Theme part",
        ));
    }
    let replaced_snapshot = read_theme_family_part_value(&replaced)?;
    let replaced_family = replaced_snapshot
        .family()
        .ok_or_else(|| fail("theme-family replace lost its direct owner"))?;
    let requested_name = changed_family.name().to_owned();
    let requested_id = changed_family.id().to_string();
    let requested_variant_id = changed_family.variant_id().to_string();
    let replaced_identity_matches = replaced_family == &changed_family;
    if !replaced_identity_matches {
        return Err(fail(
            "theme-family replace parsed owner differs from requested identity",
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
    let added_family = added_snapshot
        .family()
        .ok_or_else(|| fail("theme-family add did not restore a direct owner"))?;
    let added_identity_matches = added_family == &changed_family;
    if !added_identity_matches {
        return Err(fail(
            "theme-family add parsed owner differs from requested identity",
        ));
    }
    Ok(vec![
        ThemeFamilyMutationReceipt {
            operation: "replace_family",
            source_sha256: source_sha256.clone(),
            output_sha256: sha256(&replaced),
            output_bytes: replaced.len(),
            output_differs: true,
            requested_name: Some(requested_name.clone()),
            requested_id: Some(requested_id.clone()),
            requested_variant_id: Some(requested_variant_id.clone()),
            observed_name: Some(replaced_family.name().to_owned()),
            observed_id: Some(replaced_family.id().to_string()),
            observed_variant_id: Some(replaced_family.variant_id().to_string()),
            requested_identity_matches: replaced_identity_matches,
            source_name: source_name.clone(),
            source_id: source_id.clone(),
            source_variant_id: source_variant_id.clone(),
            source_identity_matches,
        },
        ThemeFamilyMutationReceipt {
            operation: "remove_family",
            source_sha256: source_sha256.clone(),
            output_sha256: sha256(&removed),
            output_bytes: removed.len(),
            output_differs: removed != part.source,
            requested_name: None,
            requested_id: None,
            requested_variant_id: None,
            observed_name: None,
            observed_id: None,
            observed_variant_id: None,
            requested_identity_matches: removed_snapshot.family().is_none(),
            source_name: source_name.clone(),
            source_id: source_id.clone(),
            source_variant_id: source_variant_id.clone(),
            source_identity_matches,
        },
        ThemeFamilyMutationReceipt {
            operation: "add_family",
            source_sha256: sha256(&removed),
            output_sha256: sha256(&added),
            output_bytes: added.len(),
            output_differs: added != removed,
            requested_name: Some(requested_name),
            requested_id: Some(requested_id),
            requested_variant_id: Some(requested_variant_id),
            observed_name: Some(added_family.name().to_owned()),
            observed_id: Some(added_family.id().to_string()),
            observed_variant_id: Some(added_family.variant_id().to_string()),
            requested_identity_matches: added_identity_matches,
            source_name,
            source_id,
            source_variant_id,
            source_identity_matches,
        },
    ])
}

fn preflight_check(case: &LaneCase) -> Result<serde_json::Value, ProbeFailure> {
    let mut parts = Vec::with_capacity(case.parts.len());
    let mut total_sink_writes = 0usize;
    let mut total_partial_output = 0usize;
    let mut all_partial_outputs_within_limit = true;
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
        let prepared = case
            .prepared
            .iter()
            .find(|value| value.id == part.id)
            .ok_or_else(|| fail(format!("missing prepared value for {}", part.id)))?;
        if !output_limit_supported(part.kind) {
            parts.push(serde_json::json!({
                "part_id": part.id,
                "boundary": boundary,
                "production_refusal": production_refusal,
                "sink_supported": supports_production_sink(part.kind),
                "preflight": "not_applicable",
                "reason": "this source endpoint has no caller-selected output cap",
                "sink_writes": 0,
            }));
            continue;
        }
        let preflight_bytes = prepared
            .output_payload_bytes
            .ok_or_else(|| fail(format!("{} output fixture has no prepared size", part.id)))?;
        let output_limit = prepared
            .output_limit
            .ok_or_else(|| fail(format!("{} output fixture has no prepared cap", part.id)))?;
        let attempt = output_limit_attempt(part, prepared, output_limit)?;
        let sink_writes = attempt.sink.as_ref().map_or(0, |sink| sink.writes);
        let partial_output_bytes = attempt.sink.as_ref().map_or(0, |sink| sink.bytes);
        total_sink_writes = total_sink_writes.saturating_add(sink_writes);
        total_partial_output = total_partial_output.saturating_add(partial_output_bytes);
        all_partial_outputs_within_limit &= partial_output_bytes <= output_limit;
        parts.push(serde_json::json!({
            "part_id": part.id,
            "boundary": boundary,
            "production_refusal": production_refusal,
            "sink_supported": supports_production_sink(part.kind),
            "preflight": "typed_output_limit",
            "preflight_bytes": preflight_bytes,
            "sink_writes": sink_writes,
            "output_limit": output_limit,
            "observed_output": attempt.refusal.observed,
            "sink_refusal": attempt.refusal,
            "partial_output_allowed": attempt.partial_output_allowed,
            "partial_output_bytes": partial_output_bytes,
            "partial_output_within_limit": partial_output_bytes <= output_limit,
            "operation_scope": match part.kind {
                PartKind::Diagram | PartKind::ThemeFamilyPart => "source_edit_serialize",
                PartKind::SvgBlip => "contextual_typed_writer",
                PartKind::TextBody | PartKind::Geometry => "string_output",
                PartKind::Theme | PartKind::Color => "vec_output",
                _ => "typed_writer",
            },
        }));
    }
    Ok(serde_json::json!({
        "preflight_zero_writes": total_sink_writes == 0,
        "preflight_writes_within_limit": all_partial_outputs_within_limit,
        "preflight_partial_output_bytes": total_partial_output,
        "sink_writes": total_sink_writes,
        "parts": parts,
    }))
}

/// Execute only production boundary calls during a timed preflight row.
/// Fixture parsing, contextual SVG preparation, cap selection, and output
/// sizing are all completed in `make_case`; the timed call is only the real
/// production result.
fn production_preflight(case: &LaneCase) -> Result<(), ProbeFailure> {
    for part in &case.parts {
        if boundary_kind(part.kind) == "none" {
            continue;
        }
        let prepared = case.prepared.iter().find(|prepared| prepared.id == part.id);
        timed_entry_read(
            part.kind,
            &part.over,
            prepared.and_then(|prepared| prepared.contextual_namespace.as_ref()),
        );
    }
    for (part, prepared) in case.parts.iter().zip(&case.prepared) {
        let Some(output_limit) = prepared.output_limit else {
            continue;
        };
        let mut sink = CountingSink::timed(usize::MAX);
        output_limit_timed(part, prepared, output_limit, &mut sink);
    }
    Ok(())
}

/// The timed malformed row calls each real reader/writer route.  It does not
/// serialize a JSON receipt or calculate a digest until the outside-clock
/// detail pass.
fn malformed_production(case: &LaneCase) -> Result<(), ProbeFailure> {
    for (part, prepared) in case.parts.iter().zip(&case.prepared) {
        timed_entry_read(
            part.kind,
            &part.malformed,
            prepared.contextual_namespace.as_ref(),
        );
    }
    Ok(())
}

fn normal_write(case: &LaneCase, timed: bool) -> Result<OperationValue, ProbeFailure> {
    if !case
        .parts
        .iter()
        .any(|part| supports_production_sink(part.kind))
    {
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
    if timed {
        let mut sink = CountingSink::timed(usize::MAX);
        for (part, prepared) in case.parts.iter().zip(&case.prepared) {
            if !supports_production_sink(part.kind) {
                continue;
            }
            if part.kind == PartKind::SvgBlip {
                if let Some(value) = prepared.contextual_changed.as_ref() {
                    #[cfg(feature = "drawingml-raw-current")]
                    black_box(dml::svg_blip::write_contextual_to(
                        &mut sink,
                        value,
                        dml::svg_blip::Limits::DEFAULT,
                    ));
                    #[cfg(not(feature = "drawingml-raw-current"))]
                    black_box(dml::svg_blip::write_contextual_to(
                        &mut sink,
                        value,
                        dml::svg_blip::MAX_XML_BYTES,
                    ));
                }
            } else {
                timed_write_prepared_to(part.kind, &prepared.changed, &mut sink);
            }
        }
        Ok(OperationValue::Timed)
    } else {
        let mut sink = CountingSink::new(usize::MAX);
        for (part, prepared) in case.parts.iter().zip(&case.prepared) {
            if !supports_production_sink(part.kind) {
                continue;
            }
            if part.kind == PartKind::SvgBlip {
                let value = prepared
                    .contextual_changed
                    .as_ref()
                    .ok_or_else(|| fail("SVG contextual value was not prepared"))?;
                write_svg_contextual_to_default(&mut sink, value)?;
            } else {
                write_prepared_to(part.kind, &prepared.changed, &mut sink)?;
            }
        }
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
                "contextual_svg_parts": case
                    .parts
                    .iter()
                    .filter(|part| part.kind == PartKind::SvgBlip)
                    .map(|part| part.id)
                    .collect::<Vec<_>>(),
                "contextual_svg_scope": "contextual_read_write",
            }),
        })
    }
}

/// Call the raw readers directly for the timed entry row.  Each arm leaves
/// its native production `Result` opaque to the optimizer; no common probe
/// error or refusal receipt is built on the clock.
fn timed_entry_read(
    kind: PartKind,
    bytes: &[u8],
    contextual_namespace: Option<&dml::svg_blip::NamespaceContext>,
) {
    match kind {
        PartKind::Chart => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::chart::reader::read(
                bytes,
                &dml::chart::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::chart::reader::read(bytes));
        },
        PartKind::FormatCode => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::chart::extension::formatcode2::read(
                bytes,
                &dml::chart::extension::formatcode2::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::chart::extension::formatcode2::read(bytes));
        },
        PartKind::Diagram => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::diagram::edit::Outline::read(
                bytes,
                &dml::diagram::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::diagram::edit::Outline::read(bytes));
        },
        PartKind::TextBody => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::text::body::reader::read(
                bytes,
                quick_xml::name::NamespaceResolver::default(),
                dml::text::body::reader::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::text::body::reader::read_with_options(
                bytes,
                quick_xml::name::NamespaceResolver::default(),
                dml::text::body::reader::Limits::DEFAULT,
            ));
        },
        PartKind::SvgBlip => {
            if let Some(context) = contextual_namespace {
                #[cfg(feature = "drawingml-raw-current")]
                black_box(dml::svg_blip::read_contextual(
                    bytes,
                    context,
                    dml::svg_blip::Limits::DEFAULT,
                ));
                #[cfg(not(feature = "drawingml-raw-current"))]
                black_box(dml::svg_blip::read_contextual(bytes, context));
            }
        },
        PartKind::Model3d => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::model3d::codec::read(
                bytes,
                dml::model3d::codec::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::model3d::codec::read(bytes));
        },
        PartKind::Blip => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::blip::find_first_embed(
                bytes,
                dml::blip::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::blip::find_first_embed(bytes));
        },
        PartKind::Theme => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::theme::codec::read(
                bytes,
                dml::theme::codec::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::theme::codec::read(bytes));
        },
        PartKind::ThemeFamily => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::theme::family::read(
                bytes,
                dml::theme::family::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::theme::family::read(bytes));
        },
        PartKind::ThemeFamilyPart => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::theme::family::part::read(
                bytes,
                dml::theme::family::part::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::theme::family::part::read(bytes));
        },
        PartKind::InkActions => {
            #[cfg(any(
                feature = "drawingml-raw-current",
                feature = "drawingml-ink-limits-current"
            ))]
            black_box(dml::ink::actions::read_profile(
                bytes,
                dml::ink::actions::Limits::DEFAULT,
            ));
            #[cfg(not(any(
                feature = "drawingml-raw-current",
                feature = "drawingml-ink-limits-current"
            )))]
            black_box(dml::ink::actions::read_profile(bytes));
        },
        PartKind::Geometry => {
            // Keep the native geometry Result opaque.  `parse_geometry` is
            // the untimed receipt adapter; normalizing it to ProbeFailure in
            // this row would measure harness error construction instead of
            // the production builder walk.
            black_box(parse_geometry_native(bytes));
        },
        PartKind::Color => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::color::read(bytes, &dml::color::Limits::DEFAULT));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::color::read(bytes));
        },
        PartKind::Transform => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::transform::codec::read(
                bytes,
                &dml::transform::codec::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::transform::codec::read(bytes));
        },
    }
}

fn entry_read_operation(case: &LaneCase, timed: bool) -> Result<OperationValue, ProbeFailure> {
    if timed {
        for (part, prepared) in case.parts.iter().zip(&case.prepared) {
            timed_entry_read(
                part.kind,
                &part.source,
                prepared.contextual_namespace.as_ref(),
            );
        }
        return Ok(OperationValue::Timed);
    }
    let mut sink = CountingSink::new(usize::MAX);
    for (part, prepared) in case.parts.iter().zip(&case.prepared) {
        let output = write_prepared_vec(part.kind, &prepared.source)?;
        sink.write_all(&output)
            .map_err(|error| io_failure(error, output.len()))?;
    }
    Ok(OperationValue::Sink(write_receipt(sink)))
}

/// Timed Vec/String writes mirror `write_prepared_vec` without routing native
/// results through `ProbeFailure`.  Source-edit routes still call their real
/// production mutation and serializer; all receipt/fidelity work remains in
/// the pre-warmup detail pass.
fn timed_write_prepared_vec(kind: PartKind, value: &PreparedValue) {
    match (kind, value) {
        (PartKind::Chart, PreparedValue::Chart(value)) => {
            let mut output = Vec::new();
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::chart::writer::write(
                &mut output,
                value,
                &dml::chart::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::chart::writer::write(&mut output, value));
        },
        (PartKind::FormatCode, PreparedValue::FormatCode(value)) => {
            let mut output = Vec::new();
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::chart::extension::formatcode2::write_to(
                &mut output,
                value,
                &dml::chart::extension::formatcode2::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::chart::extension::formatcode2::write_to(
                &mut output,
                value,
            ));
        },
        (
            PartKind::Diagram,
            PreparedValue::Diagram {
                source,
                entry_id,
                entry_text,
                drawing_relationship,
                ..
            },
        ) => {
            #[cfg(feature = "drawingml-raw-current")]
            {
                let limits = dml::diagram::Limits::DEFAULT;
                let result = dml::diagram::edit::set_text(source, entry_id, entry_text, &limits);
                let result = result.and_then(|output| {
                    dml::diagram::edit::set_drawing_reference(
                        &output,
                        drawing_relationship.as_deref(),
                        &limits,
                    )
                });
                black_box(result);
            }
            #[cfg(not(feature = "drawingml-raw-current"))]
            {
                let result = dml::diagram::edit::set_text(source, entry_id, entry_text);
                let result = result.and_then(|output| {
                    dml::diagram::edit::set_drawing_reference(
                        &output,
                        drawing_relationship.as_deref(),
                    )
                });
                black_box(result);
            }
        },
        (PartKind::TextBody, PreparedValue::TextBody(value)) => {
            let mut contents = String::new();
            #[cfg(feature = "drawingml-raw-current")]
            let result = dml::text::body::writer::write(
                &mut contents,
                value,
                dml::text::body::writer::Limits::DEFAULT,
            );
            #[cfg(not(feature = "drawingml-raw-current"))]
            let result = dml::text::body::writer::write(&mut contents, value);
            // Keep the production writer's result and bytes observable without
            // adding the host wrapper serialization to the timed route.
            black_box((result, contents));
        },
        (PartKind::Model3d, PreparedValue::Model3d(value)) => {
            let mut output = Vec::new();
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::model3d::codec::write_to(
                &mut output,
                value,
                dml::model3d::codec::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::model3d::codec::write_to(&mut output, value));
        },
        (PartKind::Blip, PreparedValue::Blip(value)) => {
            let mut output = String::new();
            black_box(dml::blip::write_embed(&mut output, value, true));
        },
        (PartKind::Theme, PreparedValue::Theme(value)) => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::theme::codec::encode_part(
                &value.name,
                &value.colors,
                &value.fonts,
                dml::theme::codec::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::theme::codec::encode_part(
                &value.name,
                &value.colors,
                &value.fonts,
            ));
        },
        (PartKind::ThemeFamily, PreparedValue::ThemeFamily(value)) => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::theme::family::write(
                value,
                dml::theme::family::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::theme::family::write(value));
        },
        (PartKind::ThemeFamilyPart, PreparedValue::ThemeFamilyPart { snapshot, family }) => {
            #[cfg(feature = "drawingml-theme-family-current")]
            black_box(snapshot.replace_family(family, dml::theme::family::part::Limits::DEFAULT));
            #[cfg(not(feature = "drawingml-theme-family-current"))]
            black_box(
                snapshot.replace_family_with_limit(family, dml::theme::family::part::MAX_XML_BYTES),
            );
        },
        (PartKind::InkActions, PreparedValue::InkActions(value)) => {
            let mut output = Vec::new();
            #[cfg(any(
                feature = "drawingml-raw-current",
                feature = "drawingml-ink-limits-current"
            ))]
            black_box(dml::ink::actions::write_profile_to(
                &mut output,
                value,
                dml::ink::actions::Limits::DEFAULT,
            ));
            #[cfg(not(any(
                feature = "drawingml-raw-current",
                feature = "drawingml-ink-limits-current"
            )))]
            black_box(dml::ink::actions::write_profile_to(&mut output, value));
        },
        (PartKind::Geometry, PreparedValue::Geometry(value)) => {
            let mut output = String::new();
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::geometry::writer::write_custom_geometry(
                &mut output,
                value,
                &dml::geometry::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::geometry::writer::write_custom_geometry(
                &mut output,
                value,
            ));
        },
        (PartKind::Color, PreparedValue::Color(value)) => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::color::write(value, &dml::color::Limits::DEFAULT));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::color::write(value));
        },
        (PartKind::Transform, PreparedValue::Transform(value)) => {
            let mut output = Vec::new();
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::transform::codec::write_to(
                &mut output,
                value,
                &dml::transform::codec::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::transform::codec::write_to(&mut output, value));
        },
        _ => {},
    }
}

/// Direct native sink calls for the bounded-sink row.  Only the capability
/// matrix's real sink writers reach this helper.
fn timed_write_prepared_to(kind: PartKind, value: &PreparedValue, sink: &mut CountingSink) {
    match (kind, value) {
        (PartKind::Chart, PreparedValue::Chart(value)) => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::chart::writer::write(
                sink,
                value,
                &dml::chart::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::chart::writer::write(sink, value));
        },
        (PartKind::FormatCode, PreparedValue::FormatCode(value)) => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::chart::extension::formatcode2::write_to(
                sink,
                value,
                &dml::chart::extension::formatcode2::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::chart::extension::formatcode2::write_to(sink, value));
        },
        (PartKind::Model3d, PreparedValue::Model3d(value)) => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::model3d::codec::write_to(
                sink,
                value,
                dml::model3d::codec::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::model3d::codec::write_to(sink, value));
        },
        (PartKind::ThemeFamily, PreparedValue::ThemeFamily(value)) => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::theme::family::write_to(
                sink,
                value,
                dml::theme::family::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::theme::family::write_to(sink, value));
        },
        (PartKind::InkActions, PreparedValue::InkActions(value)) => {
            #[cfg(any(
                feature = "drawingml-raw-current",
                feature = "drawingml-ink-limits-current"
            ))]
            black_box(dml::ink::actions::write_profile_to(
                sink,
                value,
                dml::ink::actions::Limits::DEFAULT,
            ));
            #[cfg(not(any(
                feature = "drawingml-raw-current",
                feature = "drawingml-ink-limits-current"
            )))]
            black_box(dml::ink::actions::write_profile_to(sink, value));
        },
        (PartKind::Transform, PreparedValue::Transform(value)) => {
            #[cfg(feature = "drawingml-raw-current")]
            black_box(dml::transform::codec::write_to(
                sink,
                value,
                &dml::transform::codec::Limits::DEFAULT,
            ));
            #[cfg(not(feature = "drawingml-raw-current"))]
            black_box(dml::transform::codec::write_to(sink, value));
        },
        _ => {},
    }
}

/// Write timing consumes values parsed once in `make_case`; no reader or
/// source round trip is hidden in the writer sample. Diagram and Theme Family
/// part entries intentionally exercise their public source-edit-and-serialize
/// operations, and their receipt names that scope rather than calling it a
/// detached pure writer.
fn vec_write_operation(case: &LaneCase, timed: bool) -> Result<OperationValue, ProbeFailure> {
    if timed {
        for (part, prepared) in case.parts.iter().zip(&case.prepared) {
            if part.kind == PartKind::SvgBlip {
                if let Some(value) = prepared.contextual_changed.as_ref() {
                    #[cfg(feature = "drawingml-raw-current")]
                    black_box(dml::svg_blip::write_contextual(
                        value,
                        dml::svg_blip::Limits::DEFAULT,
                    ));
                    #[cfg(not(feature = "drawingml-raw-current"))]
                    black_box(dml::svg_blip::write_contextual(
                        value,
                        dml::svg_blip::MAX_XML_BYTES,
                    ));
                }
            } else {
                timed_write_prepared_vec(part.kind, &prepared.changed);
            }
        }
        return Ok(OperationValue::Timed);
    }
    let mut sink = CountingSink::new(usize::MAX);
    for (part, prepared) in case.parts.iter().zip(&case.prepared) {
        let output = if part.kind == PartKind::SvgBlip {
            let value = prepared
                .contextual_changed
                .as_ref()
                .ok_or_else(|| fail("SVG contextual value was not prepared"))?;
            write_svg_contextual_default(value)?
        } else {
            write_prepared_vec(part.kind, &prepared.changed)?
        };
        sink.write_all(&output)
            .map_err(|error| io_failure(error, output.len()))?;
    }
    Ok(OperationValue::SinkDetails {
        sink: write_receipt(sink),
        details: serde_json::json!({
            "operation_scope": "mixed_prepared_output",
            "source_edit_serialize_parts": case
                .parts
                .iter()
                .filter(|part| matches!(part.kind, PartKind::Diagram | PartKind::ThemeFamilyPart))
                .map(|part| part.id)
                .collect::<Vec<_>>(),
            "string_output_parts": case
                .parts
                .iter()
                .filter(|part| matches!(part.kind, PartKind::TextBody | PartKind::Geometry))
                .map(|part| part.id)
                .collect::<Vec<_>>(),
            "vec_output_parts": case
                .parts
                .iter()
                .filter(|part| matches!(part.kind, PartKind::Theme | PartKind::Color))
                .map(|part| part.id)
                .collect::<Vec<_>>(),
            "scanner_helper_parts": case
                .parts
                .iter()
                .filter(|part| part.kind == PartKind::Blip)
                .map(|part| part.id)
                .collect::<Vec<_>>(),
            "scanner_helper_scope": "scanner_helper",
            "contextual_svg_parts": case
                .parts
                .iter()
                .filter(|part| part.kind == PartKind::SvgBlip)
                .map(|part| part.id)
                .collect::<Vec<_>>(),
            "contextual_svg_scope": "contextual_read_write",
            "typed_writer_parts": case
                .parts
                .iter()
                .filter(|part| {
                    !matches!(
                        part.kind,
                        PartKind::Diagram
                            | PartKind::ThemeFamilyPart
                            | PartKind::TextBody
                            | PartKind::Geometry
                            | PartKind::Theme
                            | PartKind::Color
                            | PartKind::Blip
                    )
                })
                .map(|part| part.id)
                .collect::<Vec<_>>(),
        }),
    })
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
    let mut stale_identities = Vec::new();
    for work in &case.prepared_owners {
        let fixture = &work.fixture;
        match operation_name {
            "noop_commit" => {
                if timed {
                    black_box_native_owner(owner_native_commit(work, false));
                } else {
                    let output = owner_native_commit_output(
                        fixture.route,
                        owner_native_commit(work, false),
                    )?;
                    sink.write_all(&output)
                        .map_err(|error| io_failure(error, output.len()))?;
                }
            },
            "changed_commit" => {
                if timed {
                    black_box_native_owner(owner_native_commit(work, true));
                } else {
                    let output =
                        owner_native_commit_output(fixture.route, owner_native_commit(work, true))?;
                    sink.write_all(&output)
                        .map_err(|error| io_failure(error, output.len()))?;
                }
            },
            "forward_replay" => {
                if timed {
                    black_box_native_replay(owner_native_replay(&work.prepared, &work.patch));
                } else {
                    let output = owner_native_replay_output(
                        fixture.route,
                        owner_native_replay(&work.prepared, &work.patch),
                    )?;
                    sink.write_all(&output)
                        .map_err(|error| io_failure(error, output.len()))?;
                }
            },
            "inverse_replay" => {
                if timed {
                    black_box_native_replay(owner_native_replay(
                        &work.changed_prepared,
                        &work.inverse_patch,
                    ));
                } else {
                    let output = owner_native_replay_output(
                        fixture.route,
                        owner_native_replay(&work.changed_prepared, &work.inverse_patch),
                    )?;
                    sink.write_all(&output)
                        .map_err(|error| io_failure(error, output.len()))?;
                }
            },
            "stale_replay" => {
                if timed {
                    black_box_native_replay(owner_native_replay(&work.stale_prepared, &work.patch));
                } else {
                    let refusal = owner_native_stale_refusal(work)?;
                    let reason_verified = stale_reason_verified(fixture.route, &refusal);
                    if !reason_verified {
                        return Err(fail(format!(
                            "{} stale replay did not preserve its exact source reason",
                            owner_route_name(fixture.route)
                        )));
                    }
                    stale_identities.push(serde_json::json!({
                        "route": owner_route_name(fixture.route),
                        "error_variant": stale_error_variant(),
                        "reason_verified": reason_verified,
                    }));
                    refusals.push(refusal);
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
                "stale_error_identities": stale_identities,
                "native_operation": owner_operation_path(
                    case.prepared_owners[0].fixture.route,
                    operation_name,
                ),
                "error_checked_outside_clock": true,
            }))
        })
    } else if timed {
        Ok(OperationValue::Timed)
    } else {
        Ok(OperationValue::SinkDetails {
            sink: write_receipt(sink),
            details: serde_json::json!({
                "native_operation": owner_operation_path(
                    case.prepared_owners[0].fixture.route,
                    operation_name,
                ),
                "result_serialized_outside_clock": true,
            }),
        })
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
    prepared_value: OperationValue,
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
    let (details, sink) = match prepared_value {
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

fn write_json_value<T: Serialize>(path: &Path, value: &T) -> io::Result<()> {
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
        serde_json::to_vec_pretty(value).map_err(|error| io::Error::other(error.to_string()))?;
    let mut file = fs::File::create(&temporary)?;
    file.write_all(&bytes)?;
    file.write_all(b"\n")?;
    file.sync_all()?;
    fs::rename(temporary, path)
}

fn write_report(path: &Path, report: &Report) -> io::Result<()> {
    write_json_value(path, report)
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
    // Freeze the generated bytes before any warmup or timed row.  The paired
    // validator compares this receipt across baseline/candidate legs so a
    // source-generation drift cannot be mistaken for an endpoint speedup.
    let generated_fixture_parts = build_generated_fixture_parts(&case);
    let generated_fixture_sha256 = generated_fixture_identity(&generated_fixture_parts);
    if arguments.preflight {
        // This handoff is deliberately outside the measured protocol.  The
        // paired runner compares source/changed/malformed identities from the
        // two endpoint processes before it starts warmups or ABBA children.
        // Keep the report small and free of timing/RSS rows so it cannot be
        // mistaken for a measured child report.
        let report = serde_json::json!({
            "schema": "litchi.drawingml.raw.preflight.v1",
            "status": "preflight",
            "measured": false,
            "qualification_claim": false,
            "lane": arguments.lane,
            "endpoint": arguments.endpoint,
            "owner_route": arguments.owner_route,
            "raw_scope": arguments.raw_scope,
            "role": arguments.role,
            "mode": arguments.mode,
            "source_revision": arguments.source_revision,
            "fixture_bytes": fixture_bytes.len(),
            "fixture_sha256": sha256(&fixture_bytes),
            "fixture_lane_identity_sha256": case.fixture_identity_sha256,
            "generated_fixture_sha256": generated_fixture_sha256,
            "generated_fixture_parts": generated_fixture_parts,
        });
        return write_json_value(&arguments.output, &report);
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
        case.prepared_owners
            .iter()
            .map(|work| owner_proof(&case.fixture_identity_sha256, work))
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
    // Build every outside-clock operation receipt before the first warmup.
    // Timed rows receive these immutable details and never rerun a fidelity,
    // refusal, digest, or output-size check after their samples.
    let prepared_operations = operations
        .iter()
        .map(|operation_name| operation(&case, operation_name, false))
        .collect::<Result<Vec<_>, _>>()
        .map_err(|error| io::Error::other(error.message))?;
    let mut rows = Vec::with_capacity(operations.len());
    let mut all_metrics = Vec::new();
    let mut peak_rss = process_rss_bytes();
    for (operation_name, prepared_value) in operations.into_iter().zip(prepared_operations) {
        let (row, metrics, row_peak) = measure(
            &case,
            &operation_name,
            &arguments,
            observer,
            !raw_fidelity.is_empty() || !owner_proofs.is_empty(),
            prepared_value,
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
        endpoint: arguments.endpoint.clone(),
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
        generated_fixture_sha256,
        generated_fixture_parts,
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
