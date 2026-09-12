//! Bounded public-API adapter for the PPTX existing-target InkAction owner.
//!
//! This module owns synthetic, complete OPC graphs only.  It deliberately
//! does not claim that the graphs were produced or accepted by PowerPoint.
//! Fixture construction and all expected-value work happen before the timed
//! operation.  The runner supplies process isolation and `/usr/bin/time -v`
//! RSS evidence; this binary supplies public API calls and phase-local
//! allocation receipts.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::cast_precision_loss,
    clippy::expect_used,
    clippy::print_stdout,
    clippy::shadow_reuse,
    reason = "the bounded evidence adapter emits deterministic JSON and uses synthetic fixtures"
)]

use std::collections::{BTreeSet, HashMap};
use std::error::Error as StdError;
use std::time::Instant;

use litchi_drawingml::ink::actions::{ActionSelector, ActionType, ChildSelector};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, Part, ReadLimits, TargetMode};
use litchi_pptx::presentation::embedded::ink_actions::{Commit, Limits, Patch, Snapshot};
use litchi_pptx::{Error as PptxError, Package};
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};

use crate::support::{self, AllocSnapshot, PhaseRecord};

type BoxError = Box<dyn StdError + Send + Sync>;
type Result<T> = std::result::Result<T, BoxError>;

const SCHEMA: &str = "pptx-ink-actions-performance-v1";
const SEMANTIC_OWNER_COMMIT: &str = "cf6fdb8e91dd232d7d762596763d2e9d8a5b9dbd";
const PRODUCTION_SOURCE_BASELINE_COMMIT: &str = "2a2ffa1cae4e6b7070082768ce84483e5d411dc8";
const SEMANTIC_OWNER_DESIGN_PATH: &str =
    "docs/report/spec-gap-validation-evidence/pptx-ink-actions-design.md";
const SEMANTIC_OWNER_DESIGN_SHA256: &str =
    "30b78cca84c4ca24ae44f3d3694c3097f54b5e5a1f2004af9ce5007bcaf4173d";
const SEMANTIC_OWNER_DESIGN_GIT_BLOB: &str = "597400950b1027c47cd6e4cbbedd23915bc0980e";
// `source_commit` remains a compatibility receipt field for the semantic
// owner. It is never the production baseline or the runtime capture HEAD.
const SOURCE_COMMIT: &str = SEMANTIC_OWNER_COMMIT;
const HELPER_SHA256: &str = "bec6baafcf735d778216fb54fe6299312707e912dcaed5f89de988f7112bb58e";

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const STRICT_CUSTOM_XML: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const P14_MAIN: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const ACTION: &str = "http://schemas.microsoft.com/office/powerpoint/2014/inkAction";
const INKML: &str = "http://www.w3.org/2003/InkML";
const SLIDE_PREFIX: &str = "/ppt/slides/slide";
const ACTION_PREFIX: &str = "/ppt/custom/action";
const OPAQUE_CHOICE_MARKER: &[u8] = b"<v:u";
const OPAQUE_FALLBACK_MARKER: &[u8] = b"<v:f";
const OPAQUE_PAYLOAD_MARKER: &[u8] = b"opaque";
const OPAQUE_DEFAULT_NAMESPACE_MARKER: &[u8] = b"defaultOpaque";
const OPAQUE_PREFIX_MARKER: &[u8] = b"prefixOpaque";
const OPAQUE_UNKNOWN_REQUIRES_MARKER: &[u8] = b"unknownInkFeature";

const MANIFEST_BYTES: &[u8] = include_bytes!("../corpus-manifest.json");
const GENERATOR_SOURCE_BYTES: &[u8] = include_bytes!("adapter.rs");

#[derive(Clone, Debug, Deserialize)]
struct Manifest {
    schema: String,
    semantic_owner_commit: String,
    production_source_baseline_commit: String,
    semantic_owner_design: SemanticOwnerDesign,
    #[serde(default)]
    owner_commit: String,
    fixture_authority: FixtureAuthority,
    retained_opc_generator: GeneratorAuthority,
    recipes: Vec<RecipeSpec>,
    lanes: Vec<LaneSpec>,
    measurement_gate: MeasurementGate,
}

#[derive(Clone, Debug, Deserialize)]
struct SemanticOwnerDesign {
    commit: String,
    path: String,
    sha256: String,
    git_blob: String,
}

#[derive(Clone, Debug, Deserialize)]
struct GeneratorAuthority {
    path: String,
    sha256: String,
    git_blob: String,
}

#[derive(Clone, Debug, Deserialize)]
struct FixtureAuthority {
    path: String,
    sha256: String,
    git_blob: String,
}

#[derive(Clone, Debug, Deserialize)]
struct MeasurementGate {
    recipe_count: usize,
    lane_count: usize,
    fresh_processes: usize,
    warmups_per_process: usize,
    samples_per_process: usize,
}

#[derive(Clone, Debug, Deserialize)]
struct RecipeSpec {
    id: String,
    anchors: usize,
    target_bytes: usize,
    unique_targets: usize,
    edges: usize,
    #[serde(default)]
    outbound_edges: usize,
    slides: usize,
    dialect: String,
    topology: String,
    #[serde(default)]
    opaque_mce: bool,
    #[serde(default)]
    unknown_internal_outbound: bool,
    #[serde(default)]
    unknown_external_outbound: bool,
    #[serde(default)]
    signed: bool,
    #[serde(default)]
    stale_input: Option<String>,
    #[serde(default)]
    limit_resource: Option<String>,
    #[serde(default)]
    limits: Vec<usize>,
}

#[derive(Clone, Debug, Deserialize)]
struct LaneSpec {
    id: String,
    recipe: String,
    operation: String,
    #[serde(default)]
    expected: Option<String>,
}

#[derive(Clone, Debug)]
#[allow(dead_code)]
struct FixtureFacts {
    source_bytes: Vec<u8>,
    source_sha256: String,
    source_fnv1a64: u64,
    package_manifest: PackageManifest,
    package_manifest_sha256: String,
    slide_names: Vec<PackURI>,
    target_names: Vec<PackURI>,
}

struct Prepared {
    lane: LaneSpec,
    recipe: RecipeSpec,
    package: Package,
    operation_package: Option<Package>,
    retained_baseline_live_bytes: u64,
    facts: FixtureFacts,
    package_before_bytes: Vec<u8>,
    package_before_manifest: PackageManifest,
    package_before_manifest_sha256: String,
    read_limits: ReadLimits,
    owner_limits: Limits,
    source_snapshot: Option<Snapshot>,
    commit: Option<Commit>,
    inverse_patch: Option<Patch>,
    operation: OperationKind,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum OperationKind {
    PackageRead,
    PresentationRead,
    PackageReadWithLimits,
    PackageEdit,
    PresentationEdit,
    PackageNoop,
    PresentationNoop,
    Apply,
    Inverse,
    Save,
    SaveReopen,
    Limit,
}

#[derive(Debug)]
enum HeldResult {
    Snapshots(Vec<Snapshot>),
    Snapshot(Snapshot),
    Commit(Commit),
    Bytes(Vec<u8>),
    SaveReopen {
        bytes: Vec<u8>,
        snapshots: Vec<Snapshot>,
    },
}

#[derive(Clone, Debug)]
struct ActualError {
    kind: String,
    debug: String,
    display: String,
    resource: Option<String>,
    limit: Option<u64>,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
struct RelationshipManifest {
    id: String,
    relationship_type: String,
    target_ref: String,
    target_mode: String,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
struct RelationshipMemberManifest {
    present: bool,
    bytes: usize,
    sha256: String,
    relationships: Vec<RelationshipManifest>,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
struct PartManifest {
    name: String,
    content_type: String,
    bytes: usize,
    sha256: String,
    relationships: RelationshipMemberManifest,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq, PartialOrd, Ord)]
struct NonPartMemberManifest {
    name: String,
    reason: String,
    bytes: usize,
    sha256: String,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
struct PackageManifest {
    content_types_bytes: usize,
    content_types_sha256: String,
    package_relationships: RelationshipMemberManifest,
    parts: Vec<PartManifest>,
    non_part_members: Vec<NonPartMemberManifest>,
}

#[derive(Debug)]
enum HeldOutcome {
    Success(HeldResult),
    Error(ActualError),
}

#[derive(Clone, Debug, Serialize)]
struct ValidationSummary {
    semantic_ok: bool,
    preservation_ok: bool,
    inverse_ok: bool,
    source_unchanged_on_refusal: bool,
    package_manifest_preserved: bool,
    expected_error: Option<String>,
    actual_error_type: Option<String>,
    actual_error_resource: Option<String>,
    actual_error_limit: Option<u64>,
    actual_error_debug: Option<String>,
    actual_error_display: Option<String>,
}

#[derive(Clone, Debug)]
struct GraphMetrics {
    source_bytes: usize,
    owner_xml_bytes: usize,
    unique_target_bytes: usize,
    baseline_unique_target_bytes: usize,
    expected_unique_target_bytes: usize,
    anchors: usize,
    unique_targets: usize,
    inbound_edges: usize,
    outbound_edges: usize,
    action_count: usize,
    action_group_count: usize,
    retained_owner_xml_bytes: usize,
    retained_target_bytes: usize,
    retained_profile_bytes: usize,
    shared_pointer_observation: String,
    same_target_pointer_consistent: bool,
    distinct_target_pointer_isolated: bool,
    outbound_diagnostic_modes: Vec<String>,
    unknown_internal_outbound_preserved: bool,
    unknown_external_outbound_preserved: bool,
    opaque_choice_preserved: bool,
    opaque_fallback_preserved: bool,
    opaque_payload_preserved: bool,
    opaque_default_namespace_preserved: bool,
    opaque_prefix_preserved: bool,
    opaque_unknown_requires_preserved: bool,
    owner_xml_sha256: String,
    profile_source_sha256: String,
}

#[derive(Clone, Debug, Default)]
struct TargetMetrics {
    target_bytes: usize,
    inbound_edges: usize,
    outbound_edges: usize,
    profile_bytes: usize,
    target_modes: Vec<TargetMode>,
    pointer: Option<usize>,
    pointer_consistent: bool,
    action_count: usize,
    action_group_count: usize,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct PointerObservation {
    same_target_pointer_consistent: bool,
    distinct_target_pointer_isolated: bool,
    label: &'static str,
}

#[derive(Clone, Debug)]
struct ValidationReport {
    summary: ValidationSummary,
    metrics: GraphMetrics,
    output_bytes: Option<usize>,
    output_manifest_sha256: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
struct SampleReceipt {
    process_id: u32,
    semantic_ok: bool,
    preservation_ok: bool,
    inverse_ok: bool,
    source_unchanged_on_refusal: bool,
    package_manifest_preserved: bool,
    expected_error: Option<String>,
    actual_error_type: Option<String>,
    actual_error_resource: Option<String>,
    actual_error_limit: Option<u64>,
    actual_error_debug: Option<String>,
    actual_error_display: Option<String>,
    source_sha256: String,
    source_fnv1a64: u64,
    source_manifest_sha256: String,
    package_before_manifest_sha256: String,
    output_manifest_sha256: Option<String>,
    source_bytes: usize,
    owner_xml_bytes: usize,
    unique_target_bytes: usize,
    baseline_unique_target_bytes: usize,
    expected_unique_target_bytes: usize,
    output_bytes: Option<usize>,
    anchors: usize,
    unique_targets: usize,
    inbound_edges: usize,
    outbound_edges: usize,
    action_count: usize,
    action_group_count: usize,
    retained_owner_xml_bytes: usize,
    retained_target_bytes: usize,
    retained_profile_bytes: usize,
    shared_pointer_observation: String,
    same_target_pointer_consistent: bool,
    distinct_target_pointer_isolated: bool,
    outbound_diagnostic_modes: Vec<String>,
    unknown_internal_outbound_preserved: bool,
    unknown_external_outbound_preserved: bool,
    opaque_choice_preserved: bool,
    opaque_fallback_preserved: bool,
    opaque_payload_preserved: bool,
    opaque_default_namespace_preserved: bool,
    opaque_prefix_preserved: bool,
    opaque_unknown_requires_preserved: bool,
    owner_xml_sha256: String,
    profile_source_sha256: String,
    retained_baseline_live_bytes: u64,
    after_drop_live_bytes: u64,
    after_postdrop_live_bytes: u64,
    baseline_release_bytes: u64,
    baseline_reopenable: bool,
    retained_baseline_balance_ok: bool,
    phases: PhaseSet,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct PhaseSet {
    setup: PhaseRecord,
    operation: PhaseRecord,
    validation: PhaseRecord,
    drop: PhaseRecord,
    postdrop: PhaseRecord,
}

/// Fixed-size receipt staging keeps validation text and metric strings out of
/// the drop/post-drop retained set.  Dynamic receipt allocations are created
/// only after the prepared baseline has been released.
#[derive(Clone, Copy)]
struct FixedText {
    length: u16,
    bytes: [u8; 512],
}

impl FixedText {
    fn from_str(value: &str) -> Self {
        let source = value.as_bytes();
        let length = source.len().min(512);
        let mut bytes = [0; 512];
        bytes[..length].copy_from_slice(&source[..length]);
        Self {
            length: length as u16,
            bytes,
        }
    }

    fn to_string(self) -> String {
        String::from_utf8_lossy(&self.bytes[..usize::from(self.length)]).into_owned()
    }
}

#[derive(Clone, Copy)]
struct DigestBytes([u8; 32]);

fn digest_bytes(value: &str) -> Result<DigestBytes> {
    if value.len() != 64 {
        return Err(format!("digest has {} characters, expected 64", value.len()).into());
    }
    let mut bytes = [0; 32];
    for (index, byte) in bytes.iter_mut().enumerate() {
        *byte = u8::from_str_radix(&value[index * 2..index * 2 + 2], 16)
            .map_err(|error| format!("invalid digest: {error}"))?;
    }
    Ok(DigestBytes(bytes))
}

fn digest_string(value: DigestBytes) -> String {
    value.0.iter().map(|byte| format!("{byte:02x}")).collect()
}

#[derive(Clone, Copy)]
struct CompactError {
    kind: FixedText,
    debug: FixedText,
    display: FixedText,
    resource: Option<FixedText>,
    limit: Option<u64>,
}

#[derive(Clone, Copy)]
struct CompactMetrics {
    source_bytes: usize,
    owner_xml_bytes: usize,
    unique_target_bytes: usize,
    baseline_unique_target_bytes: usize,
    expected_unique_target_bytes: usize,
    anchors: usize,
    unique_targets: usize,
    inbound_edges: usize,
    outbound_edges: usize,
    action_count: usize,
    action_group_count: usize,
    retained_owner_xml_bytes: usize,
    retained_target_bytes: usize,
    retained_profile_bytes: usize,
    shared_pointer_observation: FixedText,
    same_target_pointer_consistent: bool,
    distinct_target_pointer_isolated: bool,
    diagnostic_modes: u8,
    unknown_internal_outbound_preserved: bool,
    unknown_external_outbound_preserved: bool,
    opaque_choice_preserved: bool,
    opaque_fallback_preserved: bool,
    opaque_payload_preserved: bool,
    opaque_default_namespace_preserved: bool,
    opaque_prefix_preserved: bool,
    opaque_unknown_requires_preserved: bool,
    owner_xml_sha256: DigestBytes,
    profile_source_sha256: DigestBytes,
}

#[derive(Clone, Copy)]
struct CompactSample {
    semantic_ok: bool,
    preservation_ok: bool,
    inverse_ok: bool,
    source_unchanged_on_refusal: bool,
    package_manifest_preserved: bool,
    error: Option<CompactError>,
    source_sha256: DigestBytes,
    source_fnv1a64: u64,
    source_manifest_sha256: DigestBytes,
    package_before_manifest_sha256: DigestBytes,
    output_manifest_sha256: Option<DigestBytes>,
    metrics: CompactMetrics,
    output_bytes: Option<usize>,
    retained_baseline_live_bytes: u64,
    after_drop_live_bytes: u64,
    after_postdrop_live_bytes: u64,
    baseline_release_bytes: u64,
    baseline_reopenable: bool,
    retained_baseline_balance_ok: bool,
    phases: PhaseSet,
}

#[derive(Clone, Debug, Serialize)]
struct Receipt {
    schema: &'static str,
    lane: String,
    recipe_id: String,
    source_commit: &'static str,
    semantic_owner_commit: &'static str,
    production_source_baseline_commit: &'static str,
    capture_head: String,
    helper_sha256: &'static str,
    generator_sha256: Option<String>,
    fixture_path: String,
    fixture_git_blob: String,
    generator_source_sha256: String,
    source_sha256: String,
    source_fnv1a64: u64,
    source_manifest_sha256: String,
    package_before_manifest_sha256: String,
    retained_opc_limits: Value,
    owner_limits: Value,
    warmup: usize,
    sample_count: usize,
    samples: Vec<SampleReceipt>,
}

fn manifest() -> Manifest {
    serde_json::from_slice(MANIFEST_BYTES).expect("committed corpus manifest is valid JSON")
}

pub fn known_lanes() -> Vec<String> {
    manifest().lanes.into_iter().map(|lane| lane.id).collect()
}

pub fn run_lane(lane_id: &str, warmup: usize, samples: usize) -> Result<String> {
    let manifest = manifest();
    validate_manifest(&manifest)?;
    let capture_head = capture_head()?;
    let lane = manifest
        .lanes
        .iter()
        .find(|lane| lane.id == lane_id)
        .cloned()
        .ok_or_else(|| format!("unknown lane: {lane_id}"))?;
    let recipe = manifest
        .recipes
        .iter()
        .find(|recipe| recipe.id == lane.recipe)
        .cloned()
        .ok_or_else(|| format!("lane recipe is absent: {}", lane.recipe))?;
    let owner_limits = owner_limits_for(&recipe, &lane)?;
    let read_limits = ReadLimits::default();
    support::reset_counters();

    for _ in 0..warmup {
        let (prepared, setup) = prepare_with_phase(&lane, &recipe, owner_limits, read_limits)?;
        let _ = run_sample(prepared, setup, true, lane.expected.as_deref())?;
    }

    let mut receipts = Vec::with_capacity(samples);
    for _ in 0..samples {
        let (prepared, setup) = prepare_with_phase(&lane, &recipe, owner_limits, read_limits)?;
        receipts.push(run_sample(
            prepared,
            setup,
            false,
            lane.expected.as_deref(),
        )?);
    }

    let receipt = Receipt {
        schema: SCHEMA,
        lane: lane.id,
        recipe_id: recipe.id,
        source_commit: SOURCE_COMMIT,
        semantic_owner_commit: SEMANTIC_OWNER_COMMIT,
        production_source_baseline_commit: PRODUCTION_SOURCE_BASELINE_COMMIT,
        capture_head,
        helper_sha256: HELPER_SHA256,
        generator_sha256: Some(support::sha256_hex(GENERATOR_SOURCE_BYTES)),
        fixture_path: manifest.fixture_authority.path,
        fixture_git_blob: manifest.fixture_authority.git_blob,
        generator_source_sha256: support::sha256_hex(GENERATOR_SOURCE_BYTES),
        source_sha256: receipts
            .first()
            .map(|sample| sample.source_sha256.clone())
            .unwrap_or_default(),
        source_fnv1a64: receipts.first().map_or(0, |sample| sample.source_fnv1a64),
        source_manifest_sha256: receipts
            .first()
            .map(|sample| sample.source_manifest_sha256.clone())
            .unwrap_or_default(),
        package_before_manifest_sha256: receipts
            .first()
            .map(|sample| sample.package_before_manifest_sha256.clone())
            .unwrap_or_default(),
        retained_opc_limits: read_limits_json(read_limits),
        owner_limits: limits_json(owner_limits),
        warmup,
        sample_count: samples,
        samples: receipts,
    };
    serde_json::to_string(&receipt).map_err(Into::into)
}

pub fn run_matrix() -> Result<String> {
    let manifest = manifest();
    validate_manifest(&manifest)?;
    let capture_head = capture_head()?;
    let mut accepted = 0usize;
    let mut failed = Vec::new();
    let mut expected_refusals = 0usize;
    for lane in &manifest.lanes {
        let recipe = manifest
            .recipes
            .iter()
            .find(|recipe| recipe.id == lane.recipe)
            .ok_or_else(|| format!("lane recipe is absent: {}", lane.recipe))?;
        let limits = owner_limits_for(recipe, lane)?;
        let (prepared, setup) = prepare_with_phase(lane, recipe, limits, ReadLimits::default())?;
        let receipt = run_sample(prepared, setup, false, lane.expected.as_deref())?;
        if receipt.semantic_ok
            && receipt.preservation_ok
            && receipt.inverse_ok
            && receipt.baseline_reopenable
            && receipt.retained_baseline_balance_ok
        {
            accepted += 1;
            if lane
                .expected
                .as_deref()
                .is_some_and(|expected| expected != "success")
            {
                expected_refusals += 1;
            }
        } else {
            failed.push(lane.id.clone());
        }
    }
    if !failed.is_empty() {
        return Err(format!(
            "correctness matrix failed {} lane(s): {}",
            failed.len(),
            failed.join(", ")
        )
        .into());
    }
    serde_json::to_string(&json!({
        "schema": "pptx-ink-actions-correctness-matrix-v1",
        "timings_collected": false,
        "source_commit": SOURCE_COMMIT,
        "semantic_owner_commit": SEMANTIC_OWNER_COMMIT,
        "production_source_baseline_commit": PRODUCTION_SOURCE_BASELINE_COMMIT,
        "capture_head": capture_head,
        "helper_sha256": HELPER_SHA256,
        "recipe_count": manifest.recipes.len(),
        "lane_count": manifest.lanes.len(),
        "accepted_lanes": accepted,
        "expected_refusal_lanes": expected_refusals,
        "failed_lanes": failed.len(),
    }))
    .map_err(Into::into)
}

pub fn host_probe() -> Result<String> {
    let manifest = manifest();
    validate_manifest(&manifest)?;
    let capture_head = capture_head()?;
    let lane = manifest
        .lanes
        .iter()
        .find(|lane| lane.id == "package_read_tiny_shared")
        .ok_or("host probe lane is absent")?;
    let recipe = manifest
        .recipes
        .iter()
        .find(|recipe| recipe.id == lane.recipe)
        .ok_or("host probe recipe is absent")?;
    let limits = owner_limits_for(recipe, lane)?;
    let prepared = prepare(lane, recipe, limits, ReadLimits::default())?;
    let package = prepared
        .package
        .ink_actions()
        .map_err(|error| format!("package host probe failed: {error:?}"))?;
    let presentation = prepared
        .package
        .presentation()
        .and_then(|presentation| presentation.ink_actions())
        .map_err(|error| format!("presentation host probe failed: {error:?}"))?;
    let package_anchors = package
        .iter()
        .map(|snapshot| snapshot.anchors().len())
        .sum::<usize>();
    let presentation_anchors = presentation
        .iter()
        .map(|snapshot| snapshot.anchors().len())
        .sum::<usize>();
    drop(prepared);
    let opaque_lane = manifest
        .lanes
        .iter()
        .find(|lane| lane.id == "opaque_mce_scalar_edit")
        .ok_or("opaque host probe lane is absent")?;
    let opaque_recipe = manifest
        .recipes
        .iter()
        .find(|recipe| recipe.id == opaque_lane.recipe)
        .ok_or("opaque host probe recipe is absent")?;
    let opaque_limits = owner_limits_for(opaque_recipe, opaque_lane)?;
    let mut opaque_prepared = prepare(
        opaque_lane,
        opaque_recipe,
        opaque_limits,
        ReadLimits::default(),
    )?;
    let opaque_snapshots = active_package(&opaque_prepared)
        .ink_actions_with_limits(opaque_limits)
        .map_err(|error| format!("opaque host probe failed: {error:?}"))?;
    let opaque_metrics = graph_metrics(&opaque_prepared.facts, &opaque_snapshots);
    if !preservation_ok(&opaque_prepared, &opaque_metrics) {
        return Err("opaque host probe did not retain unknown MCE markers".into());
    }
    let opaque_manifest_preserved = opaque_prepared
        .package
        .to_bytes()
        .ok()
        .and_then(|bytes| package_manifest_from_bytes(&bytes, ReadLimits::default()).ok())
        .is_some_and(|(manifest, _)| manifest == opaque_prepared.package_before_manifest);
    if !opaque_manifest_preserved {
        return Err("opaque host probe did not preserve exact member manifests".into());
    }
    let opaque_source = opaque_prepared
        .source_snapshot
        .as_ref()
        .ok_or("opaque host probe source snapshot is absent")?;
    let opaque_commit = make_commit(opaque_source, true)?;
    let applied = active_package_mut(&mut opaque_prepared)
        .apply_ink_actions_patch(opaque_commit.patch())
        .map_err(|error| format!("opaque publication host probe failed: {error:?}"))?;
    drop(applied);
    let published_bytes = active_package_mut(&mut opaque_prepared)
        .to_bytes()
        .map_err(|error| format!("opaque publication serialization failed: {error:?}"))?;
    let (published_manifest, _) =
        package_manifest_from_bytes(&published_bytes, ReadLimits::default())?;
    let opaque_target_names = mutable_target_names(&opaque_prepared);
    let opaque_patch_manifest_preserved = manifest_preserves_topology(
        &opaque_prepared.package_before_manifest,
        &published_manifest,
        &opaque_target_names,
    );
    let opaque_patch_slices_preserved = opaque_target_bytes_preserved(
        &opaque_prepared.package_before_bytes,
        &published_bytes,
        &opaque_target_names,
        ReadLimits::default(),
        true,
    );
    let reopened = Package::from_vec_with_limits(published_bytes, ReadLimits::default())
        .map_err(|error| format!("opaque publication reopen failed: {error:?}"))?;
    let reopened_snapshots = reopened
        .ink_actions_with_limits(opaque_limits)
        .map_err(|error| format!("opaque publication readback failed: {error:?}"))?;
    let reopened_metrics = graph_metrics(&opaque_prepared.facts, &reopened_snapshots);
    let opaque_patch_publication_preserved = opaque_patch_manifest_preserved
        && opaque_patch_slices_preserved
        && preservation_ok(&opaque_prepared, &reopened_metrics);
    if !opaque_patch_publication_preserved {
        return Err(
            "opaque host probe did not preserve exact bytes through patch publication".into(),
        );
    }
    serde_json::to_string(&json!({
        "schema": "pptx-ink-actions-host-probe-v1",
        "source_commit": SOURCE_COMMIT,
        "semantic_owner_commit": SEMANTIC_OWNER_COMMIT,
        "production_source_baseline_commit": PRODUCTION_SOURCE_BASELINE_COMMIT,
        "capture_head": capture_head,
        "helper_sha256": HELPER_SHA256,
        "package_route": "Package::ink_actions",
        "presentation_route": "Presentation::ink_actions",
        "package_anchors": package_anchors,
        "presentation_anchors": presentation_anchors,
        "opaque_choice_preserved": opaque_metrics.opaque_choice_preserved,
        "opaque_fallback_preserved": opaque_metrics.opaque_fallback_preserved,
        "opaque_payload_preserved": opaque_metrics.opaque_payload_preserved,
        "opaque_default_namespace_preserved": opaque_metrics.opaque_default_namespace_preserved,
        "opaque_prefix_preserved": opaque_metrics.opaque_prefix_preserved,
        "opaque_unknown_requires_preserved": opaque_metrics.opaque_unknown_requires_preserved,
        "opaque_internal_outbound_preserved": opaque_metrics.unknown_internal_outbound_preserved,
        "opaque_external_outbound_preserved": opaque_metrics.unknown_external_outbound_preserved,
        "opaque_manifest_preserved": opaque_manifest_preserved,
        "opaque_patch_publication_preserved": opaque_patch_publication_preserved,
        "native_powerpoint_claim": false,
        "synthetic_complete_opc": true,
    }))
    .map_err(Into::into)
}

fn validate_manifest(manifest: &Manifest) -> Result<()> {
    if manifest.schema != "pptx-ink-actions-performance-scaffold-v2" {
        return Err(format!("unexpected manifest schema: {}", manifest.schema).into());
    }
    if manifest.semantic_owner_commit != SEMANTIC_OWNER_COMMIT {
        return Err(format!(
            "manifest semantic owner pin differs: {}",
            manifest.semantic_owner_commit
        )
        .into());
    }
    if manifest.production_source_baseline_commit != PRODUCTION_SOURCE_BASELINE_COMMIT {
        return Err(format!(
            "manifest production source baseline differs: {}",
            manifest.production_source_baseline_commit
        )
        .into());
    }
    if manifest.owner_commit != SEMANTIC_OWNER_COMMIT {
        return Err(format!("manifest owner alias differs: {}", manifest.owner_commit).into());
    }
    if manifest.semantic_owner_design.commit != SEMANTIC_OWNER_COMMIT
        || manifest.semantic_owner_design.path != SEMANTIC_OWNER_DESIGN_PATH
        || manifest.semantic_owner_design.sha256 != SEMANTIC_OWNER_DESIGN_SHA256
        || manifest.semantic_owner_design.git_blob != SEMANTIC_OWNER_DESIGN_GIT_BLOB
    {
        return Err("manifest semantic owner design authority differs".into());
    }
    if manifest.fixture_authority.sha256 != HELPER_SHA256 {
        return Err("manifest helper hash differs from compiled provenance".into());
    }
    if manifest.retained_opc_generator.path != "harness/adapter.rs"
        || manifest.retained_opc_generator.sha256 != support::sha256_hex(GENERATOR_SOURCE_BYTES)
        || manifest.retained_opc_generator.git_blob.len() != 40
    {
        return Err("manifest retained generator hash differs from compiled source".into());
    }
    if manifest.measurement_gate.recipe_count != manifest.recipes.len()
        || manifest.measurement_gate.lane_count != manifest.lanes.len()
    {
        return Err("manifest count fields do not match recipe/lane arrays".into());
    }
    if manifest.measurement_gate.fresh_processes != 3
        || manifest.measurement_gate.warmups_per_process != 2
        || manifest.measurement_gate.samples_per_process != 20
    {
        return Err("manifest measurement gate changed".into());
    }
    let recipe_ids = manifest
        .recipes
        .iter()
        .map(|recipe| recipe.id.as_str())
        .collect::<BTreeSet<_>>();
    if recipe_ids.len() != manifest.recipes.len() {
        return Err("recipe IDs are not unique".into());
    }
    let lane_ids = manifest
        .lanes
        .iter()
        .map(|lane| lane.id.as_str())
        .collect::<BTreeSet<_>>();
    if lane_ids.len() != manifest.lanes.len() {
        return Err("lane IDs are not unique".into());
    }
    let usage = manifest
        .lanes
        .iter()
        .map(|lane| lane.recipe.as_str())
        .collect::<BTreeSet<_>>();
    if usage != recipe_ids {
        return Err(format!(
            "recipe usage mismatch: orphaned={:?}, unresolved={:?}",
            recipe_ids.difference(&usage).collect::<Vec<_>>(),
            usage.difference(&recipe_ids).collect::<Vec<_>>()
        )
        .into());
    }
    if manifest.lanes.iter().any(|lane| lane.operation.is_empty()) {
        return Err("lane operation text is empty".into());
    }
    Ok(())
}

fn capture_head() -> Result<String> {
    let value = std::env::var("PPTX_INK_ACTIONS_CAPTURE_HEAD")
        .map_err(|_| "PPTX_INK_ACTIONS_CAPTURE_HEAD is not set")?;
    if value.len() != 40 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err("PPTX_INK_ACTIONS_CAPTURE_HEAD is not a full commit id".into());
    }
    Ok(value)
}

fn operation_for(lane: &LaneSpec) -> Result<OperationKind> {
    let id = lane.id.as_str();
    let operation = if id.starts_with("package_read_multislide") {
        OperationKind::PackageReadWithLimits
    } else if id.starts_with("package_read_") {
        OperationKind::PackageRead
    } else if id == "presentation_read_small_shared" {
        OperationKind::PresentationRead
    } else if id.starts_with("package_scalar_edit") {
        OperationKind::PackageEdit
    } else if id.starts_with("presentation_scalar_edit") {
        OperationKind::PresentationEdit
    } else if id.starts_with("package_noop") {
        OperationKind::PackageNoop
    } else if id.starts_with("presentation_noop") {
        OperationKind::PresentationNoop
    } else if id.starts_with("package_apply")
        || id.starts_with("stale_")
        || id.starts_with("signed_")
        || id == "opaque_mce_scalar_edit"
    {
        OperationKind::Apply
    } else if id.starts_with("package_inverse") {
        OperationKind::Inverse
    } else if id == "package_save_medium_shared" {
        OperationKind::Save
    } else if id == "package_save_reopen_medium_shared" || id == "opaque_mce_save_reopen" {
        OperationKind::SaveReopen
    } else if id.starts_with("limit_") {
        OperationKind::Limit
    } else if id == "unknown_outbound_read_edit" || id == "strict_shared_edit" {
        OperationKind::PackageEdit
    } else {
        return Err(format!("lane operation is not wired: {id}").into());
    };
    Ok(operation)
}

fn prepare(
    lane: &LaneSpec,
    recipe: &RecipeSpec,
    owner_limits: Limits,
    read_limits: ReadLimits,
) -> Result<Prepared> {
    let operation = operation_for(lane)?;
    // These values are retained in `Prepared` through the operation and named
    // drop phase. Clone them before the retained-baseline snapshot so their
    // setup allocations are charged to the retained prepared set. The apply
    // working package is intentionally constructed after that snapshot and
    // released in `drop`.
    let prepared_lane = lane.clone();
    let prepared_recipe = recipe.clone();
    let (_, baseline_facts) = build_opc(recipe)?;
    let source_snapshot = if matches!(
        operation,
        OperationKind::PackageEdit
            | OperationKind::PresentationEdit
            | OperationKind::PackageNoop
            | OperationKind::PresentationNoop
            | OperationKind::Apply
            | OperationKind::Inverse
    ) {
        let source_package = Package::from_opc_package(build_opc(recipe)?.0)
            .map_err(|error| format!("source package setup failed: {error:?}"))?;
        let snapshots = if matches!(
            operation,
            OperationKind::PresentationEdit | OperationKind::PresentationNoop
        ) {
            source_package
                .presentation()
                .map_err(|error| format!("source presentation setup failed: {error:?}"))?
                .ink_actions_with_limits(owner_limits)
        } else {
            source_package.ink_actions_with_limits(owner_limits)
        }
        .map_err(|error| format!("source snapshot setup failed: {error:?}"))?;
        snapshots.into_iter().next()
    } else {
        None
    };

    let commit = if matches!(operation, OperationKind::Apply | OperationKind::Inverse) {
        let source = source_snapshot
            .as_ref()
            .ok_or("patch lane has no source snapshot")?;
        Some(make_commit(source, lane.id != "signed_noop")?)
    } else {
        None
    };
    let inverse_patch = commit.as_ref().map(|commit| commit.patch().inverse());

    let target_opc = target_opc_for_recipe(recipe)?;
    let mut package = Package::from_opc_package(target_opc)
        .map_err(|error| format!("public mutable package setup failed: {error:?}"))?;
    let package_before_bytes = package
        .to_bytes()
        .map_err(|error| format!("package baseline serialization failed: {error:?}"))?;
    let (package_before_manifest, package_before_manifest_sha256) =
        package_manifest_from_bytes(&package_before_bytes, read_limits)?;
    let retained_baseline_live_bytes = AllocSnapshot::now().live_bytes();
    let operation_package = if matches!(operation, OperationKind::Apply | OperationKind::Inverse) {
        let operation_opc = target_opc_for_recipe(recipe)?;
        Some(
            Package::from_opc_package(operation_opc)
                .map_err(|error| format!("apply operation package setup failed: {error:?}"))?,
        )
    } else {
        None
    };
    Ok(Prepared {
        lane: prepared_lane,
        recipe: prepared_recipe,
        package,
        operation_package,
        retained_baseline_live_bytes,
        facts: baseline_facts,
        package_before_bytes,
        package_before_manifest,
        package_before_manifest_sha256,
        read_limits,
        owner_limits,
        source_snapshot,
        commit,
        inverse_patch,
        operation,
    })
}

fn target_opc_for_recipe(recipe: &RecipeSpec) -> Result<OpcPackage> {
    if recipe.stale_input.is_some() {
        let (mutated, _) = build_opc(recipe)?;
        let mut mutated = mutated;
        mutate_stale(&mut mutated, recipe.stale_input.as_deref().unwrap())?;
        Ok(mutated)
    } else {
        Ok(build_opc(recipe)?.0)
    }
}

fn prepare_with_phase(
    lane: &LaneSpec,
    recipe: &RecipeSpec,
    owner_limits: Limits,
    read_limits: ReadLimits,
) -> Result<(Prepared, PhaseRecord)> {
    let before = AllocSnapshot::phase_start();
    let started = Instant::now();
    let prepared = prepare(lane, recipe, owner_limits, read_limits)?;
    let after = AllocSnapshot::now();
    Ok((prepared, before.delta(after, started.elapsed())))
}

fn owner_limits_for(recipe: &RecipeSpec, lane: &LaneSpec) -> Result<Limits> {
    if recipe.limit_resource.is_none() {
        return Ok(Limits::default());
    }
    let resource = recipe.limit_resource.as_deref().unwrap();
    let index = if lane.id.ends_with("one_under") {
        0
    } else if lane.id.ends_with("exact") {
        1
    } else if lane.id.ends_with("one_over") {
        2
    } else {
        return Err(format!("limit lane lacks a boundary suffix: {}", lane.id).into());
    };
    let value = *recipe
        .limits
        .get(index)
        .ok_or_else(|| format!("limit recipe has no boundary {index}: {}", recipe.id))?;
    let defaults = Limits::default();
    let values = match resource {
        "anchors" => (
            value,
            defaults.target_bytes,
            defaults.total_target_bytes,
            defaults.target_relationships,
        ),
        "target_bytes" => (
            defaults.anchors,
            value,
            defaults.total_target_bytes,
            defaults.target_relationships,
        ),
        "total_target_bytes" => (
            defaults.anchors,
            defaults.target_bytes,
            value,
            defaults.target_relationships,
        ),
        "target_relationships" => (
            defaults.anchors,
            defaults.target_bytes,
            defaults.total_target_bytes,
            value,
        ),
        unknown => return Err(format!("unknown limit resource: {unknown}").into()),
    };
    Limits::new(values.0, values.1, values.2, values.3)
        .ok_or_else(|| format!("invalid owner limits for {resource}: {value}").into())
}

fn make_commit(source: &Snapshot, changed: bool) -> Result<Commit> {
    let mut edit = source.edit();
    if changed {
        let selector = source
            .selector(0)
            .ok_or("source snapshot has no editable anchor")?;
        edit.edit_profile(selector, |profile| {
            profile.set_action_type(ActionSelector::ordinal(0), ActionType::Transform)?;
            profile.set_start_time(ActionSelector::ordinal(0), "2.50")?;
            profile.set_property_value(
                ChildSelector::Property {
                    action: ActionSelector::ordinal(0),
                    index: 0,
                },
                "changed",
            )?;
            Ok(())
        })
        .map_err(|error| format!("profile edit setup failed: {error:?}"))?;
    }
    edit.commit()
        .map_err(|error| format!("profile commit setup failed: {error:?}").into())
}

fn build_opc(recipe: &RecipeSpec) -> Result<(OpcPackage, FixtureFacts)> {
    if recipe.slides == 0 || recipe.anchors == 0 || recipe.unique_targets == 0 {
        return Err(format!("recipe has a zero shape: {}", recipe.id).into());
    }
    let strict = recipe.dialect == "strict";
    let mut package = OpcPackage::new();
    let target_names = (0..recipe.unique_targets)
        .map(|index| {
            let suffix = if index == 0 {
                String::new()
            } else {
                index.to_string()
            };
            PackURI::new(format!("{ACTION_PREFIX}{suffix}.xml"))
                .map_err(|error| format!("target URI failed: {error:?}"))
        })
        .collect::<std::result::Result<Vec<_>, String>>()?;

    for target_name in &target_names {
        let target_payload = action_payload(recipe.target_bytes, recipe.opaque_mce)?;
        let target = BlobPart::new(target_name.clone(), "text/xml".to_owned(), target_payload);
        package.add_part(Box::new(target));
    }

    let mut slide_names = Vec::with_capacity(recipe.slides);
    let anchors_per_slide = distribute(recipe.anchors, recipe.slides);
    let mut global_anchor = 0usize;
    for slide_index in 0..recipe.slides {
        let slide_name = PackURI::new(format!("{SLIDE_PREFIX}{}.xml", slide_index + 1))
            .map_err(|error| format!("slide URI failed: {error:?}"))?;
        let count = anchors_per_slide[slide_index];
        let mut relationship_ids = Vec::with_capacity(count);
        let mut target_refs = Vec::with_capacity(count);
        for _ in 0..count {
            let relationship_id = format!("rIdAction{}", global_anchor + 1);
            relationship_ids.push(relationship_id);
            let target_index = target_index(recipe, global_anchor);
            let target_ref = if recipe.topology == "case_equivalent_shared" && global_anchor == 1 {
                "../custom/ACTION.XML".to_owned()
            } else if target_index == 0 {
                "../custom/action.xml".to_owned()
            } else {
                format!("../custom/action{target_index}.xml")
            };
            target_refs.push(target_ref);
            global_anchor += 1;
        }
        let mut slide = BlobPart::new(
            slide_name.clone(),
            ct::PML_SLIDE.to_owned(),
            slide_xml(strict, &relationship_ids, recipe.opaque_mce),
        );
        for (index, relationship_id) in relationship_ids.iter().enumerate() {
            let relationship_type = if strict {
                STRICT_CUSTOM_XML
            } else {
                rt::CUSTOM_XML
            };
            slide.rels_mut().add_relationship(
                relationship_type.to_owned(),
                target_refs[index].clone(),
                relationship_id.clone(),
                false,
            );
        }
        package.add_part(Box::new(slide));
        slide_names.push(slide_name);
    }

    let presentation_name = PackURI::new("/ppt/presentation.xml")
        .map_err(|error| format!("presentation URI failed: {error:?}"))?;
    let pml = if strict { STRICT_PML } else { PML };
    let rel = if strict { STRICT_REL } else { REL };
    let slides = slide_names
        .iter()
        .enumerate()
        .map(|(index, _)| {
            format!(
                r#"<p:sldId id="{}" r:id="rIdSlide{}"/>"#,
                256 + index,
                index + 1
            )
        })
        .collect::<String>();
    let mut presentation = BlobPart::new(
        presentation_name,
        ct::PML_PRESENTATION_MAIN.to_owned(),
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:p="{pml}" xmlns:r="{rel}"><p:sldIdLst>{slides}</p:sldIdLst></p:presentation>"#
        )
        .into_bytes(),
    );
    for (index, _) in slide_names.iter().enumerate() {
        presentation.rels_mut().add_relationship(
            rt::SLIDE.to_owned(),
            format!("slides/slide{}.xml", index + 1),
            format!("rIdSlide{}", index + 1),
            false,
        );
    }
    package.add_part(Box::new(presentation));
    package.rels_mut().add_relationship(
        rt::OFFICE_DOCUMENT.to_owned(),
        "ppt/presentation.xml".to_owned(),
        "rIdOfficeDocument".to_owned(),
        false,
    );

    if recipe.unknown_internal_outbound || recipe.unknown_external_outbound {
        let target = target_names
            .first()
            .ok_or("opaque recipe has no action target")?
            .clone();
        if recipe.unknown_internal_outbound {
            package
                .get_part_mut(&target)
                .map_err(|error| format!("opaque action part missing: {error:?}"))?
                .rels_mut()
                .add_relationship(
                    "urn:vendor:opaque-action-dependency".to_owned(),
                    "opaque-dependency.bin".to_owned(),
                    "rIdOpaqueInternal".to_owned(),
                    false,
                );
            package.add_part(Box::new(BlobPart::new(
                PackURI::new("/ppt/custom/opaque-dependency.bin")
                    .map_err(|error| format!("opaque URI failed: {error:?}"))?,
                "application/octet-stream".to_owned(),
                b"opaque-internal-target".to_vec(),
            )));
        }
        if recipe.unknown_external_outbound {
            package
                .get_part_mut(&target)
                .map_err(|error| format!("opaque action part missing: {error:?}"))?
                .rels_mut()
                .add_relationship(
                    "urn:vendor:opaque-action-external".to_owned(),
                    "https://example.invalid/opaque.xml".to_owned(),
                    "rIdOpaqueExternal".to_owned(),
                    true,
                );
        }
    }

    if recipe.signed {
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    }

    let source_bytes = serialize_opc(&package)?;
    let source_sha256 = support::sha256_hex(&source_bytes);
    let source_fnv1a64 = support::fnv1a64(&source_bytes);
    let package_manifest = package_manifest(&package, &source_bytes)?;
    let package_manifest_sha256 = package_manifest_sha256(&package_manifest)?;
    Ok((
        package,
        FixtureFacts {
            source_bytes,
            source_sha256,
            source_fnv1a64,
            package_manifest,
            package_manifest_sha256,
            slide_names,
            target_names,
        },
    ))
}

fn distribute(total: usize, groups: usize) -> Vec<usize> {
    let base = total / groups;
    let remainder = total % groups;
    (0..groups)
        .map(|index| base + usize::from(index < remainder))
        .collect()
}

fn target_index(recipe: &RecipeSpec, anchor: usize) -> usize {
    if recipe.topology == "shared" || recipe.topology == "case_equivalent_shared" {
        0
    } else {
        anchor % recipe.unique_targets
    }
}

fn action_payload(target_bytes: usize, opaque_mce: bool) -> Result<Vec<u8>> {
    let prefix = if opaque_mce {
        format!(
            r###"<?xml version="1.0" encoding="UTF-8"?><ia:actions xmlns:ia="{ACTION}" xmlns:i="{INKML}" xmlns:v="urn:vendor" xmlns:mc="{MCE}" xmlns:zz="urn:vendor-prefix" lengthUnit="cm" timeUnit="ms" xml:id="root"><i:definitions><v:future v:flag="keep"/><zz:prefixOpaque/><zz:defaultOpaque xmlns="urn:vendor-default"><defaultNode/></zz:defaultOpaque></i:definitions><ia:action xml:id="a0" type="add" startTime="0"><ia:property name="kind" value="old"/><ia:actionData xml:id="d0" name="stroke" ref="#d1"><ia:transform matrix="1,0,0,1"/><i:trace><v:opaque>opaque<mc:AlternateContent><mc:Choice Requires="v unknownInkFeature"><v:u/><zz:unknownNode/></mc:Choice><mc:Fallback><v:f/></mc:Fallback></mc:AlternateContent>"###
        )
    } else {
        format!(
            r###"<?xml version="1.0" encoding="UTF-8"?><ia:actions xmlns:ia="{ACTION}" xmlns:i="{INKML}" xmlns:v="urn:vendor" xmlns:mc="{MCE}" lengthUnit="cm" timeUnit="ms" xml:id="root"><i:definitions><v:future v:flag="keep"/></i:definitions><ia:action xml:id="a0" type="add" startTime="0"><ia:property name="kind" value="old"/><ia:actionData xml:id="d0" name="stroke" ref="#d1"><ia:transform matrix="1,0,0,1"/><i:trace><v:opaque>opaque<mc:AlternateContent><mc:Choice Requires="v"><v:u/></mc:Choice><mc:Fallback><v:f/></mc:Fallback></mc:AlternateContent>"###
        )
    };
    let suffix = r###"</v:opaque></i:trace><i:traceView/></ia:actionData><ia:actionDataGroup xml:id="dg0" name="group"><ia:actionData xml:id="d1" name="other"/></ia:actionDataGroup></ia:action><ia:actionGroup xml:id="ag0" type="transform" startTime="1"><ia:action xml:id="a1" type="remove" startTime="2"><ia:actionData/></ia:action></ia:actionGroup></ia:actions>"###;
    if target_bytes < prefix.len() + suffix.len() {
        return Err(format!("target recipe is below valid payload minimum: {target_bytes}").into());
    }
    let filler = xml_filler(target_bytes - prefix.len() - suffix.len());
    let mut output = Vec::with_capacity(target_bytes);
    output.extend_from_slice(prefix.as_bytes());
    output.extend_from_slice(&filler);
    output.extend_from_slice(suffix.as_bytes());
    debug_assert_eq!(output.len(), target_bytes);
    Ok(output)
}

fn xml_filler(length: usize) -> Vec<u8> {
    let open = b"<v:chunk>";
    let close = b"</v:chunk>";
    let overhead = open.len() + close.len();
    let mut output = Vec::with_capacity(length);
    if length <= 4096 {
        output.resize(length, b'x');
        return output;
    }
    while output.len().saturating_add(overhead).saturating_add(1) <= length {
        output.extend_from_slice(open);
        let remaining = length - output.len() - close.len();
        let chunk = remaining.min(1024);
        output.resize(output.len() + chunk, b'x');
        output.extend_from_slice(close);
    }
    output.resize(length, b'x');
    output
}

fn action_owner_anchor(relationship_id: &str, opaque_mce: bool) -> String {
    if opaque_mce {
        format!(
            r#"<compat:AlternateContent><compat:Choice Requires="p14main inkAction"><pp:contentPart rr:id="{relationship_id}"/></compat:Choice><compat:Fallback><pp:pic><pp:nvPicPr><pp:cNvPr id="42" name="Fallback"/><pp:cNvPicPr/><pp:nvPr/></pp:nvPicPr><pp:blipFill/><pp:spPr/></pp:pic></compat:Fallback></compat:AlternateContent>"#
        )
    } else {
        format!(
            r#"<compat:AlternateContent><compat:Choice Requires="p14main inkAction"><pp:contentPart rr:id="{relationship_id}"/></compat:Choice><compat:Fallback><pp:pic><pp:nvPicPr><pp:cNvPr id="42" name="Fallback"/><pp:cNvPicPr/><pp:nvPr/></pp:nvPicPr><pp:blipFill/><pp:spPr/></pp:pic></compat:Fallback></compat:AlternateContent>"#
        )
    }
}

fn slide_xml(strict: bool, relationship_ids: &[String], opaque_mce: bool) -> Vec<u8> {
    let anchors = relationship_ids
        .iter()
        .map(|relationship_id| action_owner_anchor(relationship_id, opaque_mce))
        .collect::<String>();
    let pml = if strict { STRICT_PML } else { PML };
    let rel = if strict { STRICT_REL } else { REL };
    if opaque_mce {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:p="{pml}" xmlns:q="{rel}" xmlns:pp="{pml}" xmlns:rr="{rel}" xmlns:compat="{MCE}" xmlns:p14main="{P14_MAIN}" xmlns:inkAction="{ACTION}" xmlns:unknownInkFeature="urn:vendor:unknown" xmlns:mceAlias="{MCE}" mceAlias:Ignorable="p14main inkAction unknownInkFeature"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{anchors}</p:spTree></p:cSld></p:sld>"#
        )
        .into_bytes()
    } else {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><pp:sld xmlns:pp="{pml}" xmlns:rr="{rel}" xmlns:compat="{MCE}" xmlns:p14main="{P14_MAIN}" xmlns:inkAction="{ACTION}" compat:Ignorable="p14main inkAction"><pp:cSld><pp:spTree><pp:nvGrpSpPr/><pp:grpSpPr/>{anchors}</pp:spTree></pp:cSld></pp:sld>"#
        )
        .into_bytes()
    }
}

fn serialize_opc(package: &OpcPackage) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    package
        .to_stream(&mut bytes)
        .map_err(|error| format!("OPC serialization failed: {error:?}"))?;
    Ok(bytes)
}

fn relationship_manifest(
    package: &OpcPackage,
    owner: &PackURI,
    relationships: &litchi_opc::Relationships,
) -> Result<RelationshipMemberManifest> {
    let source = package
        .source_relationships(owner)
        .map_err(|error| format!("source relationship manifest failed for {owner}: {error:?}"))?;
    let mut entries = relationships
        .iter()
        .map(|relationship| RelationshipManifest {
            id: relationship.r_id().to_owned(),
            relationship_type: relationship.reltype().to_owned(),
            target_ref: relationship.target_ref().to_owned(),
            target_mode: match relationship.target_mode() {
                TargetMode::Internal => "Internal".to_owned(),
                TargetMode::External => "External".to_owned(),
            },
        })
        .collect::<Vec<_>>();
    entries.sort_by(|left, right| left.id.cmp(&right.id));
    Ok(RelationshipMemberManifest {
        present: source.member_present(),
        bytes: source.bytes().len(),
        sha256: support::sha256_hex(source.bytes()),
        relationships: entries,
    })
}

fn package_manifest(package: &OpcPackage, archive_bytes: &[u8]) -> Result<PackageManifest> {
    let content_types = package
        .source_content_types()
        .map_err(|error| format!("source content-type manifest failed: {error:?}"))?;
    let root = PackURI::new("/").map_err(|error| format!("package URI failed: {error:?}"))?;
    let package_relationships = relationship_manifest(package, &root, package.rels())?;
    let mut parts = package
        .iter_parts()
        .map(|part| {
            Ok(PartManifest {
                name: part.partname().as_str().to_owned(),
                content_type: part.content_type().to_owned(),
                bytes: part.blob().len(),
                sha256: support::sha256_hex(part.blob()),
                relationships: relationship_manifest(package, part.partname(), part.rels())?,
            })
        })
        .collect::<Result<Vec<_>>>()?;
    parts.sort_by(|left, right| left.name.cmp(&right.name));
    let archive = soapberry_zip::office::LazyArchiveReader::new(archive_bytes)
        .map_err(|error| format!("non-part ZIP payload index failed: {error:?}"))?;
    let mut non_part_members = package
        .non_part_members()
        .iter()
        .map(|member| {
            let payload = archive.read(member.name()).map_err(|error| {
                format!(
                    "non-part ZIP payload read failed for {}: {error:?}",
                    member.name()
                )
            })?;
            Ok(NonPartMemberManifest {
                name: member.name().to_owned(),
                reason: member.reason().to_string(),
                bytes: payload.len(),
                sha256: support::sha256_hex(&payload),
            })
        })
        .collect::<Result<Vec<_>>>()?;
    non_part_members.sort();
    Ok(PackageManifest {
        content_types_bytes: content_types.bytes().len(),
        content_types_sha256: support::sha256_hex(content_types.bytes()),
        package_relationships,
        parts,
        non_part_members,
    })
}

fn package_manifest_sha256(manifest: &PackageManifest) -> Result<String> {
    let bytes = serde_json::to_vec(manifest)
        .map_err(|error| format!("package manifest serialization failed: {error}"))?;
    Ok(support::sha256_hex(&bytes))
}

fn package_manifest_from_bytes(
    bytes: &[u8],
    read_limits: ReadLimits,
) -> Result<(PackageManifest, String)> {
    let package = OpcPackage::from_bytes_with_limits(bytes, read_limits)
        .map_err(|error| format!("output OPC manifest reopen failed: {error:?}"))?;
    let manifest = package_manifest(&package, bytes)?;
    let digest = package_manifest_sha256(&manifest)?;
    Ok((manifest, digest))
}

fn package_part_bytes(bytes: &[u8], name: &PackURI, read_limits: ReadLimits) -> Result<Vec<u8>> {
    let package = OpcPackage::from_bytes_with_limits(bytes, read_limits)
        .map_err(|error| format!("part-preservation OPC reopen failed: {error:?}"))?;
    let part = package
        .get_part(name)
        .map_err(|error| format!("part-preservation member missing {name}: {error:?}"))?;
    Ok(part.blob().to_vec())
}

fn opaque_region(bytes: &[u8]) -> Option<&[u8]> {
    let start_marker = b"<v:opaque>";
    let end_marker = b"</v:opaque>";
    let start = bytes
        .windows(start_marker.len())
        .position(|window| window == start_marker)?;
    let end_start = start + start_marker.len();
    let relative_end = bytes[end_start..]
        .windows(end_marker.len())
        .position(|window| window == end_marker)?;
    let end = end_start + relative_end + end_marker.len();
    Some(&bytes[start..end])
}

fn opaque_target_bytes_preserved(
    before_bytes: &[u8],
    after_bytes: &[u8],
    target_names: &[PackURI],
    read_limits: ReadLimits,
    require_opaque: bool,
) -> bool {
    target_names.iter().all(|name| {
        let before = package_part_bytes(before_bytes, name, read_limits);
        let after = package_part_bytes(after_bytes, name, read_limits);
        match (before, after) {
            (Ok(before), Ok(after)) if before == after => true,
            (Ok(before), Ok(after)) if require_opaque => {
                opaque_region(&before).is_some() && opaque_region(&before) == opaque_region(&after)
            },
            (Ok(_), Ok(_)) => true,
            _ => false,
        }
    })
}

fn mutable_target_names(prepared: &Prepared) -> Vec<PackURI> {
    prepared
        .facts
        .target_names
        .first()
        .cloned()
        .into_iter()
        .collect()
}

fn manifest_preserves_topology(
    before: &PackageManifest,
    after: &PackageManifest,
    mutable_parts: &[PackURI],
) -> bool {
    if before.content_types_bytes != after.content_types_bytes
        || before.content_types_sha256 != after.content_types_sha256
        || before.package_relationships != after.package_relationships
        || before.non_part_members != after.non_part_members
        || before.parts.len() != after.parts.len()
    {
        return false;
    }
    before.parts.iter().zip(&after.parts).all(|(left, right)| {
        let mutable = mutable_parts.iter().any(|part| part.as_str() == left.name);
        left.name == right.name
            && left.content_type == right.content_type
            && left.relationships == right.relationships
            && (mutable || (left.bytes == right.bytes && left.sha256 == right.sha256))
    })
}

fn mutate_stale(package: &mut OpcPackage, input: &str) -> Result<()> {
    let slide_name = PackURI::new("/ppt/slides/slide1.xml")
        .map_err(|error| format!("slide URI failed: {error:?}"))?;
    let target_name = PackURI::new("/ppt/custom/action.xml")
        .map_err(|error| format!("target URI failed: {error:?}"))?;
    match input {
        "owner_xml" => {
            let part = package
                .get_part_mut(&slide_name)
                .map_err(|error| format!("stale owner part missing: {error:?}"))?;
            let mut bytes = part.blob().to_vec();
            let candidates: &[(&[u8], &[u8])] = &[
                (br#"id="42""#, br#"id="43""#),
                (br#"name="Fallback""#, br#"name="FallbacK""#),
            ];
            let (index, marker, replacement) = candidates
                .iter()
                .copied()
                .find_map(|(marker, replacement)| {
                    bytes
                        .windows(marker.len())
                        .position(|window| window == marker)
                        .map(|index| (index, marker, replacement))
                })
                .ok_or("stale owner XML marker is absent")?;
            bytes[index..index + marker.len()].copy_from_slice(replacement);
            part.set_blob(bytes);
        },
        "owner_relationship_source" => {
            package
                .get_part_mut(&slide_name)
                .map_err(|error| format!("stale owner part missing: {error:?}"))?
                .rels_mut()
                .add_relationship(
                    "urn:vendor:stale-owner-source".to_owned(),
                    "https://example.invalid/stale-owner".to_owned(),
                    "rIdStaleOwner".to_owned(),
                    true,
                );
        },
        "target_bytes" => {
            let part = package
                .get_part_mut(&target_name)
                .map_err(|error| format!("stale target part missing: {error:?}"))?;
            let marker = br#"value="old""#;
            let replacement = br#"value="other""#;
            let mut bytes = part.blob().to_vec();
            if let Some(index) = bytes
                .windows(marker.len())
                .position(|window| window == marker)
            {
                bytes.splice(index..index + marker.len(), replacement.iter().copied());
            } else {
                return Err("stale target marker is absent".into());
            }
            part.set_blob(bytes);
        },
        "content_type_source" => {
            package
                .get_part_mut(&target_name)
                .map_err(|error| format!("stale target part missing: {error:?}"))?
                .set_content_type("application/xml".to_owned())
                .map_err(|error| format!("content type mutation failed: {error:?}"))?;
        },
        unknown => return Err(format!("unknown stale recipe: {unknown}").into()),
    }
    Ok(())
}

fn empty_phase() -> PhaseRecord {
    let snapshot = AllocSnapshot::now();
    snapshot.delta(snapshot, std::time::Duration::ZERO)
}

fn compact_error(summary: &ValidationSummary) -> Option<CompactError> {
    Some(CompactError {
        kind: FixedText::from_str(summary.actual_error_type.as_deref()?),
        debug: FixedText::from_str(summary.actual_error_debug.as_deref().unwrap_or_default()),
        display: FixedText::from_str(summary.actual_error_display.as_deref().unwrap_or_default()),
        resource: summary
            .actual_error_resource
            .as_deref()
            .map(FixedText::from_str),
        limit: summary.actual_error_limit,
    })
}

fn diagnostic_bits(metrics: &GraphMetrics) -> u8 {
    metrics
        .outbound_diagnostic_modes
        .iter()
        .fold(0, |bits, mode| {
            bits | match mode.as_str() {
                "internal_unknown" => 1,
                "external_unknown" => 2,
                _ => 0,
            }
        })
}

fn compact_metrics(metrics: &GraphMetrics) -> Result<CompactMetrics> {
    Ok(CompactMetrics {
        source_bytes: metrics.source_bytes,
        owner_xml_bytes: metrics.owner_xml_bytes,
        unique_target_bytes: metrics.unique_target_bytes,
        baseline_unique_target_bytes: metrics.baseline_unique_target_bytes,
        expected_unique_target_bytes: metrics.expected_unique_target_bytes,
        anchors: metrics.anchors,
        unique_targets: metrics.unique_targets,
        inbound_edges: metrics.inbound_edges,
        outbound_edges: metrics.outbound_edges,
        action_count: metrics.action_count,
        action_group_count: metrics.action_group_count,
        retained_owner_xml_bytes: metrics.retained_owner_xml_bytes,
        retained_target_bytes: metrics.retained_target_bytes,
        retained_profile_bytes: metrics.retained_profile_bytes,
        shared_pointer_observation: FixedText::from_str(&metrics.shared_pointer_observation),
        same_target_pointer_consistent: metrics.same_target_pointer_consistent,
        distinct_target_pointer_isolated: metrics.distinct_target_pointer_isolated,
        diagnostic_modes: diagnostic_bits(metrics),
        unknown_internal_outbound_preserved: metrics.unknown_internal_outbound_preserved,
        unknown_external_outbound_preserved: metrics.unknown_external_outbound_preserved,
        opaque_choice_preserved: metrics.opaque_choice_preserved,
        opaque_fallback_preserved: metrics.opaque_fallback_preserved,
        opaque_payload_preserved: metrics.opaque_payload_preserved,
        opaque_default_namespace_preserved: metrics.opaque_default_namespace_preserved,
        opaque_prefix_preserved: metrics.opaque_prefix_preserved,
        opaque_unknown_requires_preserved: metrics.opaque_unknown_requires_preserved,
        owner_xml_sha256: digest_bytes(&metrics.owner_xml_sha256)?,
        profile_source_sha256: digest_bytes(&metrics.profile_source_sha256)?,
    })
}

fn baseline_reopenable(prepared: &Prepared) -> bool {
    let bytes = prepared.facts.source_bytes.clone();
    let mut reopened = match Package::from_vec_with_limits(bytes, prepared.read_limits) {
        Ok(package) => package,
        Err(_) => return false,
    };
    // The retained baseline is the valid, untouched source profile.  Negative
    // limit lanes deliberately use lower operation limits and must not make
    // this independent baseline probe fail.
    let snapshots = match reopened.ink_actions_with_limits(Limits::default()) {
        Ok(snapshots) => snapshots,
        Err(_) => return false,
    };
    let anchors = snapshots
        .iter()
        .map(|snapshot| snapshot.anchors().len())
        .sum::<usize>();
    if anchors != prepared.recipe.anchors {
        return false;
    }
    let serialized = match reopened.to_bytes() {
        Ok(bytes) => bytes,
        Err(_) => return false,
    };
    package_manifest_from_bytes(&serialized, prepared.read_limits)
        .map(|(manifest, _)| manifest == prepared.facts.package_manifest)
        .unwrap_or(false)
}

fn diagnostic_modes(bits: u8) -> Vec<String> {
    let mut modes = Vec::with_capacity(2);
    if bits & 1 != 0 {
        modes.push("internal_unknown".to_owned());
    }
    if bits & 2 != 0 {
        modes.push("external_unknown".to_owned());
    }
    modes
}

fn receipt_from_compact(compact: CompactSample, expected_error: Option<&str>) -> SampleReceipt {
    let metrics = compact.metrics;
    let (
        actual_error_type,
        actual_error_resource,
        actual_error_limit,
        actual_error_debug,
        actual_error_display,
    ) = match compact.error {
        Some(error) => (
            Some(error.kind.to_string()),
            error.resource.map(FixedText::to_string),
            error.limit,
            Some(error.debug.to_string()),
            Some(error.display.to_string()),
        ),
        None => (None, None, None, None, None),
    };
    SampleReceipt {
        process_id: std::process::id(),
        semantic_ok: compact.semantic_ok,
        preservation_ok: compact.preservation_ok,
        inverse_ok: compact.inverse_ok,
        source_unchanged_on_refusal: compact.source_unchanged_on_refusal,
        package_manifest_preserved: compact.package_manifest_preserved,
        expected_error: expected_error.map(str::to_owned),
        actual_error_type,
        actual_error_resource,
        actual_error_limit,
        actual_error_debug,
        actual_error_display,
        source_sha256: digest_string(compact.source_sha256),
        source_fnv1a64: compact.source_fnv1a64,
        source_manifest_sha256: digest_string(compact.source_manifest_sha256),
        package_before_manifest_sha256: digest_string(compact.package_before_manifest_sha256),
        output_manifest_sha256: compact.output_manifest_sha256.map(digest_string),
        source_bytes: metrics.source_bytes,
        owner_xml_bytes: metrics.owner_xml_bytes,
        unique_target_bytes: metrics.unique_target_bytes,
        baseline_unique_target_bytes: metrics.baseline_unique_target_bytes,
        expected_unique_target_bytes: metrics.expected_unique_target_bytes,
        output_bytes: compact.output_bytes,
        anchors: metrics.anchors,
        unique_targets: metrics.unique_targets,
        inbound_edges: metrics.inbound_edges,
        outbound_edges: metrics.outbound_edges,
        action_count: metrics.action_count,
        action_group_count: metrics.action_group_count,
        retained_owner_xml_bytes: metrics.retained_owner_xml_bytes,
        retained_target_bytes: metrics.retained_target_bytes,
        retained_profile_bytes: metrics.retained_profile_bytes,
        shared_pointer_observation: metrics.shared_pointer_observation.to_string(),
        same_target_pointer_consistent: metrics.same_target_pointer_consistent,
        distinct_target_pointer_isolated: metrics.distinct_target_pointer_isolated,
        outbound_diagnostic_modes: diagnostic_modes(metrics.diagnostic_modes),
        unknown_internal_outbound_preserved: metrics.unknown_internal_outbound_preserved,
        unknown_external_outbound_preserved: metrics.unknown_external_outbound_preserved,
        opaque_choice_preserved: metrics.opaque_choice_preserved,
        opaque_fallback_preserved: metrics.opaque_fallback_preserved,
        opaque_payload_preserved: metrics.opaque_payload_preserved,
        opaque_default_namespace_preserved: metrics.opaque_default_namespace_preserved,
        opaque_prefix_preserved: metrics.opaque_prefix_preserved,
        opaque_unknown_requires_preserved: metrics.opaque_unknown_requires_preserved,
        owner_xml_sha256: digest_string(metrics.owner_xml_sha256),
        profile_source_sha256: digest_string(metrics.profile_source_sha256),
        retained_baseline_live_bytes: compact.retained_baseline_live_bytes,
        after_drop_live_bytes: compact.after_drop_live_bytes,
        after_postdrop_live_bytes: compact.after_postdrop_live_bytes,
        baseline_release_bytes: compact.baseline_release_bytes,
        baseline_reopenable: compact.baseline_reopenable,
        retained_baseline_balance_ok: compact.retained_baseline_balance_ok,
        phases: compact.phases,
    }
}

fn run_sample(
    mut prepared: Prepared,
    setup: PhaseRecord,
    _warmup: bool,
    expected_error: Option<&str>,
) -> Result<SampleReceipt> {
    let operation_before = AllocSnapshot::phase_start();
    let operation_started = Instant::now();
    let outcome = execute_operation(&mut prepared)?;
    let operation_after = AllocSnapshot::now();
    let operation_phase = operation_before.delta(operation_after, operation_started.elapsed());

    let validation_before = AllocSnapshot::phase_start();
    let validation_started = Instant::now();
    let report = validate_outcome(&mut prepared, &outcome)?;

    let compact_metrics = compact_metrics(&report.metrics)?;
    let compact_error = compact_error(&report.summary);
    let compact_source_sha256 = digest_bytes(&prepared.facts.source_sha256)?;
    let compact_source_manifest_sha256 = digest_bytes(&prepared.facts.package_manifest_sha256)?;
    let compact_package_before_manifest_sha256 =
        digest_bytes(&prepared.package_before_manifest_sha256)?;
    let retained_baseline_live_bytes = prepared.retained_baseline_live_bytes;
    let output_manifest_sha256 = report
        .output_manifest_sha256
        .as_deref()
        .map(digest_bytes)
        .transpose()?;
    let mut compact = CompactSample {
        semantic_ok: report.summary.semantic_ok,
        preservation_ok: report.summary.preservation_ok,
        inverse_ok: report.summary.inverse_ok,
        source_unchanged_on_refusal: report.summary.source_unchanged_on_refusal,
        package_manifest_preserved: report.summary.package_manifest_preserved,
        error: compact_error,
        source_sha256: compact_source_sha256,
        source_fnv1a64: prepared.facts.source_fnv1a64,
        source_manifest_sha256: compact_source_manifest_sha256,
        package_before_manifest_sha256: compact_package_before_manifest_sha256,
        output_manifest_sha256,
        metrics: compact_metrics,
        output_bytes: report.output_bytes,
        retained_baseline_live_bytes,
        after_drop_live_bytes: 0,
        after_postdrop_live_bytes: 0,
        baseline_release_bytes: 0,
        baseline_reopenable: false,
        retained_baseline_balance_ok: false,
        phases: PhaseSet {
            setup,
            operation: operation_phase,
            validation: empty_phase(),
            drop: empty_phase(),
            postdrop: empty_phase(),
        },
    };
    let validation_after = AllocSnapshot::now();
    compact.phases.validation =
        validation_before.delta(validation_after, validation_started.elapsed());
    let drop_before = AllocSnapshot::phase_start();
    let drop_started = Instant::now();
    drop(report);
    drop(outcome);
    drop(prepared.operation_package.take());
    let drop_after = AllocSnapshot::now();
    let drop_phase = drop_before.delta(drop_after, drop_started.elapsed());

    let postdrop_before = AllocSnapshot::phase_start();
    let postdrop_started = Instant::now();
    let baseline_reopenable = baseline_reopenable(&prepared);
    drop(prepared);
    let postdrop_after = AllocSnapshot::now();
    let postdrop_phase = postdrop_before.delta(postdrop_after, postdrop_started.elapsed());
    let after_drop_live_bytes = drop_phase.live_after_bytes;
    let after_postdrop_live_bytes = postdrop_phase.live_after_bytes;
    let compact = CompactSample {
        after_drop_live_bytes,
        after_postdrop_live_bytes,
        baseline_release_bytes: after_drop_live_bytes.saturating_sub(after_postdrop_live_bytes),
        baseline_reopenable,
        retained_baseline_balance_ok: after_drop_live_bytes == retained_baseline_live_bytes,
        phases: PhaseSet {
            drop: drop_phase,
            postdrop: postdrop_phase,
            ..compact.phases
        },
        ..compact
    };
    Ok(receipt_from_compact(compact, expected_error))
}

fn active_package(prepared: &Prepared) -> &Package {
    prepared
        .operation_package
        .as_ref()
        .unwrap_or(&prepared.package)
}

fn active_package_mut(prepared: &mut Prepared) -> &mut Package {
    prepared
        .operation_package
        .as_mut()
        .unwrap_or(&mut prepared.package)
}

fn execute_operation(prepared: &mut Prepared) -> Result<HeldOutcome> {
    let outcome = match prepared.operation {
        OperationKind::PackageRead => public_snapshots(active_package(prepared).ink_actions()),
        OperationKind::PresentationRead => public_snapshots(
            active_package(prepared)
                .presentation()
                .and_then(|presentation| presentation.ink_actions()),
        ),
        OperationKind::PackageReadWithLimits => public_snapshots(
            active_package(prepared).ink_actions_with_limits(prepared.owner_limits),
        ),
        OperationKind::PackageEdit
        | OperationKind::PresentationEdit
        | OperationKind::PackageNoop
        | OperationKind::PresentationNoop => {
            let source = prepared
                .source_snapshot
                .as_ref()
                .ok_or("edit operation has no source snapshot")?;
            let mut edit = source.edit();
            let changed = matches!(
                prepared.operation,
                OperationKind::PackageEdit | OperationKind::PresentationEdit
            );
            if changed {
                let selector = source
                    .selector(0)
                    .ok_or("edit operation has no selected anchor")?;
                edit.edit_profile(selector, |profile| {
                    profile.set_action_type(ActionSelector::ordinal(0), ActionType::Transform)?;
                    profile.set_start_time(ActionSelector::ordinal(0), "2.50")?;
                    profile.set_property_value(
                        ChildSelector::Property {
                            action: ActionSelector::ordinal(0),
                            index: 0,
                        },
                        "changed",
                    )?;
                    Ok(())
                })
                .map_err(|error| format!("timed profile edit failed: {error:?}"))?;
            }
            match edit.commit() {
                Ok(commit) => HeldOutcome::Success(HeldResult::Commit(commit)),
                Err(error) => HeldOutcome::Error(actual_error(&error)),
            }
        },
        OperationKind::Apply => {
            let patch = prepared
                .commit
                .as_ref()
                .ok_or("apply operation has no prepared patch")?
                .patch()
                .clone();
            public_snapshot(active_package_mut(prepared).apply_ink_actions_patch(&patch))
        },
        OperationKind::Inverse => {
            let patch = prepared
                .commit
                .as_ref()
                .ok_or("inverse operation has no prepared patch")?
                .patch()
                .clone();
            let inverse = prepared
                .inverse_patch
                .as_ref()
                .ok_or("inverse operation has no inverse patch")?
                .clone();
            match active_package_mut(prepared).apply_ink_actions_patch(&patch) {
                Ok(_) => {
                    public_snapshot(active_package_mut(prepared).apply_ink_actions_patch(&inverse))
                },
                Err(error) => HeldOutcome::Error(actual_error(&error)),
            }
        },
        OperationKind::Save => match active_package_mut(prepared).to_bytes() {
            Ok(bytes) => HeldOutcome::Success(HeldResult::Bytes(bytes)),
            Err(error) => HeldOutcome::Error(actual_error(&error)),
        },
        OperationKind::SaveReopen => match active_package_mut(prepared).to_bytes() {
            Ok(bytes) => {
                let reopened = Package::from_vec_with_limits(bytes.clone(), prepared.read_limits);
                match reopened {
                    Ok(reopened) => match reopened.ink_actions_with_limits(prepared.owner_limits) {
                        Ok(snapshots) => {
                            HeldOutcome::Success(HeldResult::SaveReopen { bytes, snapshots })
                        },
                        Err(error) => HeldOutcome::Error(actual_error(&error)),
                    },
                    Err(error) => HeldOutcome::Error(actual_error(&error)),
                }
            },
            Err(error) => HeldOutcome::Error(actual_error(&error)),
        },
        OperationKind::Limit => public_snapshots(
            active_package(prepared).ink_actions_with_limits(prepared.owner_limits),
        ),
    };
    Ok(outcome)
}

fn public_snapshots(result: std::result::Result<Vec<Snapshot>, PptxError>) -> HeldOutcome {
    match result {
        Ok(snapshots) => HeldOutcome::Success(HeldResult::Snapshots(snapshots)),
        Err(error) => HeldOutcome::Error(actual_error(&error)),
    }
}

fn public_snapshot(result: std::result::Result<Snapshot, PptxError>) -> HeldOutcome {
    match result {
        Ok(snapshot) => HeldOutcome::Success(HeldResult::Snapshot(snapshot)),
        Err(error) => HeldOutcome::Error(actual_error(&error)),
    }
}

fn actual_error(error: &PptxError) -> ActualError {
    let debug = format!("{error:?}");
    let display = error.to_string();
    let (kind, resource, limit) = match error {
        PptxError::Limit { resource, limit } => (
            "Error::Limit".to_owned(),
            Some((*resource).to_owned()),
            Some(u64::try_from(*limit).unwrap_or(u64::MAX)),
        ),
        PptxError::ContentType { .. } => ("Error::ContentType".to_owned(), None, None),
        PptxError::StaleSource => ("Error::StaleSource".to_owned(), None, None),
        PptxError::Opc(opc_error) => match opc_error {
            OpcError::SignedSourceRequiresExplicitPolicy => (
                "Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy)".to_owned(),
                None,
                None,
            ),
            OpcError::ReadLimit {
                resource, maximum, ..
            } => (
                "Error::Opc(OpcError::ReadLimit)".to_owned(),
                Some(resource.to_string()),
                Some(*maximum),
            ),
            _ => ("Error::Opc".to_owned(), None, None),
        },
        _ => ("Error::Other".to_owned(), None, None),
    };
    ActualError {
        kind,
        debug,
        display,
        resource,
        limit,
    }
}

fn validate_outcome(prepared: &mut Prepared, outcome: &HeldOutcome) -> Result<ValidationReport> {
    let expected_error = prepared.lane.expected.clone();
    let mut summary = ValidationSummary {
        semantic_ok: false,
        preservation_ok: false,
        inverse_ok: true,
        source_unchanged_on_refusal: false,
        package_manifest_preserved: false,
        expected_error: expected_error.clone(),
        actual_error_type: None,
        actual_error_resource: None,
        actual_error_limit: None,
        actual_error_debug: None,
        actual_error_display: None,
    };

    let mut metrics = empty_metrics(&prepared.facts);
    let mut output_bytes = None;
    let mut output_manifest_sha256 = None;
    match outcome {
        HeldOutcome::Error(error) => {
            summary.actual_error_type = Some(error.kind.clone());
            summary.actual_error_resource = error.resource.clone();
            summary.actual_error_limit = error.limit;
            summary.actual_error_debug = Some(error.debug.clone());
            summary.actual_error_display = Some(error.display.clone());
            let expected = expected_error.as_deref().unwrap_or("success");
            summary.semantic_ok = expected != "success" && error_matches(expected, error);
            summary.source_unchanged_on_refusal = active_package_mut(prepared)
                .to_bytes()
                .map(|bytes| bytes == prepared.package_before_bytes)
                .unwrap_or(false);
            if let Ok(bytes) = active_package_mut(prepared).to_bytes() {
                if let Ok((manifest, digest)) =
                    package_manifest_from_bytes(&bytes, prepared.read_limits)
                {
                    summary.package_manifest_preserved =
                        manifest == prepared.package_before_manifest;
                    output_manifest_sha256 = Some(digest);
                }
            }
            summary.preservation_ok =
                summary.source_unchanged_on_refusal && summary.package_manifest_preserved;
        },
        HeldOutcome::Success(result) => {
            let snapshots = match result {
                HeldResult::Snapshots(snapshots) => {
                    if let Ok(bytes) = active_package_mut(prepared).to_bytes() {
                        if let Ok((manifest, digest)) =
                            package_manifest_from_bytes(&bytes, prepared.read_limits)
                        {
                            summary.package_manifest_preserved =
                                manifest == prepared.package_before_manifest;
                            output_manifest_sha256 = Some(digest);
                        }
                    }
                    snapshots.clone()
                },
                HeldResult::Snapshot(snapshot) => {
                    if let Ok(bytes) = active_package_mut(prepared).to_bytes() {
                        if let Ok((manifest, digest)) =
                            package_manifest_from_bytes(&bytes, prepared.read_limits)
                        {
                            summary.package_manifest_preserved =
                                if prepared.operation == OperationKind::Inverse {
                                    manifest == prepared.package_before_manifest
                                } else {
                                    manifest_preserves_topology(
                                        &prepared.package_before_manifest,
                                        &manifest,
                                        &mutable_target_names(prepared),
                                    )
                                };
                            if summary.package_manifest_preserved
                                && prepared.operation == OperationKind::Apply
                            {
                                summary.package_manifest_preserved = opaque_target_bytes_preserved(
                                    &prepared.package_before_bytes,
                                    &bytes,
                                    &mutable_target_names(prepared),
                                    prepared.read_limits,
                                    prepared.recipe.opaque_mce,
                                );
                            }
                            output_manifest_sha256 = Some(digest);
                        }
                    }
                    vec![snapshot.clone()]
                },
                HeldResult::Commit(commit) => {
                    if let Ok(bytes) = active_package_mut(prepared).to_bytes() {
                        if let Ok((manifest, digest)) =
                            package_manifest_from_bytes(&bytes, prepared.read_limits)
                        {
                            summary.package_manifest_preserved =
                                manifest == prepared.package_before_manifest;
                            output_manifest_sha256 = Some(digest);
                        }
                    }
                    vec![commit.snapshot().clone()]
                },
                HeldResult::SaveReopen { bytes, snapshots } => {
                    output_bytes = Some(bytes.len());
                    if let Ok((manifest, digest)) =
                        package_manifest_from_bytes(bytes, prepared.read_limits)
                    {
                        summary.package_manifest_preserved =
                            manifest == prepared.package_before_manifest;
                        output_manifest_sha256 = Some(digest);
                    }
                    snapshots.clone()
                },
                HeldResult::Bytes(bytes) => {
                    output_bytes = Some(bytes.len());
                    if let Ok((manifest, digest)) =
                        package_manifest_from_bytes(bytes, prepared.read_limits)
                    {
                        summary.package_manifest_preserved =
                            manifest == prepared.package_before_manifest;
                        output_manifest_sha256 = Some(digest);
                    }
                    match Package::from_vec_with_limits(bytes.clone(), prepared.read_limits)
                        .and_then(|package| package.ink_actions_with_limits(prepared.owner_limits))
                    {
                        Ok(snapshots) => snapshots,
                        Err(error) => {
                            let actual = actual_error(&error);
                            summary.actual_error_type = Some(actual.kind);
                            summary.actual_error_resource = actual.resource;
                            summary.actual_error_limit = actual.limit;
                            summary.actual_error_debug = Some(actual.debug);
                            summary.actual_error_display = Some(actual.display);
                            Vec::new()
                        },
                    }
                },
            };
            if !snapshots.is_empty() {
                metrics = graph_metrics(&prepared.facts, &snapshots);
                metrics.expected_unique_target_bytes =
                    if prepared.operation == OperationKind::Inverse {
                        prepared.recipe.target_bytes * prepared.recipe.unique_targets
                    } else {
                        snapshot_unique_target_bytes(&snapshots)
                    };
                summary.semantic_ok = semantic_shape_ok(prepared, &metrics);
                summary.preservation_ok =
                    preservation_ok(prepared, &metrics) && summary.package_manifest_preserved;
                if prepared.operation == OperationKind::Inverse {
                    summary.inverse_ok = prepared
                        .source_snapshot
                        .as_ref()
                        .is_some_and(|source| same_snapshot(source, snapshots.first().unwrap()));
                }
            }
            if expected_error
                .as_deref()
                .is_some_and(|value| value != "success")
            {
                summary.semantic_ok = false;
                summary.preservation_ok = false;
            }
            if expected_error.as_deref() == Some("success") {
                summary.semantic_ok = !snapshots.is_empty() && summary.semantic_ok;
            }
        },
    }

    Ok(ValidationReport {
        summary,
        metrics,
        output_bytes,
        output_manifest_sha256,
    })
}

fn error_matches(expected: &str, actual: &ActualError) -> bool {
    if expected == "success" {
        return false;
    }
    let expected_type = expected.split(" {").next().unwrap_or(expected);
    let type_match = actual.kind == expected_type;
    let resource_match = expected
        .split("resource:")
        .nth(1)
        .and_then(|value| value.split(',').next())
        .map(|value| {
            let expected_resource = value.trim();
            actual
                .resource
                .as_deref()
                .is_some_and(|actual_resource| actual_resource == expected_resource)
        })
        .unwrap_or(true);
    let limit_match = expected
        .split("limit:")
        .nth(1)
        .and_then(|value| value.split('}').next())
        .and_then(|value| value.trim().parse::<u64>().ok())
        .map(|expected_limit| actual.limit == Some(expected_limit))
        .unwrap_or(true);
    type_match && resource_match && limit_match
}

fn empty_metrics(facts: &FixtureFacts) -> GraphMetrics {
    GraphMetrics {
        source_bytes: facts.source_bytes.len(),
        owner_xml_bytes: 0,
        unique_target_bytes: 0,
        baseline_unique_target_bytes: 0,
        expected_unique_target_bytes: 0,
        anchors: 0,
        unique_targets: 0,
        inbound_edges: 0,
        outbound_edges: 0,
        action_count: 0,
        action_group_count: 0,
        retained_owner_xml_bytes: 0,
        retained_target_bytes: 0,
        retained_profile_bytes: 0,
        shared_pointer_observation: "none".to_owned(),
        same_target_pointer_consistent: true,
        distinct_target_pointer_isolated: true,
        outbound_diagnostic_modes: Vec::new(),
        unknown_internal_outbound_preserved: false,
        unknown_external_outbound_preserved: false,
        opaque_choice_preserved: false,
        opaque_fallback_preserved: false,
        opaque_payload_preserved: false,
        opaque_default_namespace_preserved: false,
        opaque_prefix_preserved: false,
        opaque_unknown_requires_preserved: false,
        owner_xml_sha256: support::sha256_hex(&[]),
        profile_source_sha256: support::sha256_hex(&[]),
    }
}

fn graph_metrics(facts: &FixtureFacts, snapshots: &[Snapshot]) -> GraphMetrics {
    let anchors = snapshots
        .iter()
        .flat_map(Snapshot::anchors)
        .collect::<Vec<_>>();
    let mut targets = HashMap::<String, TargetMetrics>::new();
    let mut owner_xml_bytes = 0usize;
    for anchor in &anchors {
        owner_xml_bytes = owner_xml_bytes.saturating_add(anchor.owner_xml().len());
        let target = anchor.target_part_name().as_str().to_owned();
        let target_bytes = anchor.target_bytes();
        let pointer = (!target_bytes.is_empty()).then(|| target_bytes.as_ptr() as usize);
        let entry = targets.entry(target).or_insert_with(|| TargetMetrics {
            target_bytes: target_bytes.len(),
            inbound_edges: anchor.inbound_references().len(),
            outbound_edges: anchor.outbound_references().len(),
            profile_bytes: anchor.profile().source().len(),
            target_modes: Vec::new(),
            pointer,
            pointer_consistent: pointer.is_some(),
            action_count: anchor.profile().actions().count(),
            action_group_count: anchor.profile().action_groups().count(),
        });
        entry.target_bytes = anchor.target_bytes().len();
        entry.pointer_consistent &= entry.pointer == pointer && pointer.is_some();
        for reference in anchor.outbound_references() {
            entry.target_modes.push(reference.target_mode());
        }
    }
    let pointer_observation = classify_pointer_observations(&targets);
    let mut inbound_edges = 0usize;
    let mut outbound_edges = 0usize;
    let mut unique_target_bytes = 0usize;
    let mut retained_profile_bytes = 0usize;
    let mut action_count = 0usize;
    let mut action_group_count = 0usize;
    let mut modes = BTreeSet::new();
    let mut owner_xml = Vec::new();
    let mut profile_source = Vec::new();
    for target in targets.values() {
        unique_target_bytes = unique_target_bytes.saturating_add(target.target_bytes);
        inbound_edges = inbound_edges.saturating_add(target.inbound_edges);
        outbound_edges = outbound_edges.saturating_add(target.outbound_edges);
        retained_profile_bytes = retained_profile_bytes.saturating_add(target.profile_bytes);
        action_count = action_count.saturating_add(target.action_count);
        action_group_count = action_group_count.saturating_add(target.action_group_count);
        for mode in &target.target_modes {
            modes.insert(match mode {
                TargetMode::Internal => "internal_unknown",
                TargetMode::External => "external_unknown",
            });
        }
    }
    for anchor in &anchors {
        owner_xml.extend_from_slice(anchor.owner_xml());
        profile_source.extend_from_slice(anchor.profile().source());
    }
    let opaque_choice_preserved = profile_source
        .windows(OPAQUE_CHOICE_MARKER.len())
        .any(|window| window == OPAQUE_CHOICE_MARKER);
    let opaque_fallback_preserved = profile_source
        .windows(OPAQUE_FALLBACK_MARKER.len())
        .any(|window| window == OPAQUE_FALLBACK_MARKER);
    let opaque_payload_preserved = profile_source
        .windows(OPAQUE_PAYLOAD_MARKER.len())
        .any(|window| window == OPAQUE_PAYLOAD_MARKER);
    let opaque_default_namespace_preserved = profile_source
        .windows(OPAQUE_DEFAULT_NAMESPACE_MARKER.len())
        .any(|window| window == OPAQUE_DEFAULT_NAMESPACE_MARKER);
    let opaque_prefix_preserved = profile_source
        .windows(OPAQUE_PREFIX_MARKER.len())
        .any(|window| window == OPAQUE_PREFIX_MARKER);
    let opaque_unknown_requires_preserved = profile_source
        .windows(OPAQUE_UNKNOWN_REQUIRES_MARKER.len())
        .any(|window| window == OPAQUE_UNKNOWN_REQUIRES_MARKER);
    let baseline_unique_target_bytes = facts
        .target_names
        .iter()
        .filter_map(|name| {
            facts
                .package_manifest
                .parts
                .iter()
                .find(|part| part.name == name.as_str())
                .map(|part| part.bytes)
        })
        .sum();
    GraphMetrics {
        source_bytes: facts.source_bytes.len(),
        owner_xml_bytes,
        unique_target_bytes,
        baseline_unique_target_bytes,
        expected_unique_target_bytes: baseline_unique_target_bytes,
        anchors: anchors.len(),
        unique_targets: targets.len(),
        inbound_edges,
        outbound_edges,
        action_count,
        action_group_count,
        retained_owner_xml_bytes: owner_xml_bytes,
        retained_target_bytes: unique_target_bytes,
        retained_profile_bytes,
        shared_pointer_observation: pointer_observation.label.to_owned(),
        same_target_pointer_consistent: pointer_observation.same_target_pointer_consistent,
        distinct_target_pointer_isolated: pointer_observation.distinct_target_pointer_isolated,
        outbound_diagnostic_modes: modes.into_iter().map(str::to_owned).collect(),
        unknown_internal_outbound_preserved: anchors.iter().any(|anchor| {
            anchor
                .outbound_references()
                .iter()
                .any(|reference| reference.target_mode() == TargetMode::Internal)
        }),
        unknown_external_outbound_preserved: anchors.iter().any(|anchor| {
            anchor
                .outbound_references()
                .iter()
                .any(|reference| reference.target_mode() == TargetMode::External)
        }),
        opaque_choice_preserved,
        opaque_fallback_preserved,
        opaque_payload_preserved,
        opaque_default_namespace_preserved,
        opaque_prefix_preserved,
        opaque_unknown_requires_preserved,
        owner_xml_sha256: support::sha256_hex(&owner_xml),
        profile_source_sha256: support::sha256_hex(&profile_source),
    }
}

fn classify_pointer_observations(targets: &HashMap<String, TargetMetrics>) -> PointerObservation {
    if targets.is_empty() {
        return PointerObservation {
            same_target_pointer_consistent: true,
            distinct_target_pointer_isolated: true,
            label: "none",
        };
    }

    let same_target_pointer_consistent = targets
        .values()
        .all(|target| target.pointer_consistent && target.pointer.is_some());
    let mut pointer_targets = HashMap::<usize, &str>::new();
    let mut distinct_target_pointer_isolated = true;
    for (target_name, target) in targets {
        let Some(pointer) = target.pointer else {
            distinct_target_pointer_isolated = false;
            continue;
        };
        if let Some(previous_target) = pointer_targets.insert(pointer, target_name.as_str()) {
            if previous_target != target_name {
                distinct_target_pointer_isolated = false;
            }
        }
    }
    let label = if !same_target_pointer_consistent || !distinct_target_pointer_isolated {
        "inconsistent"
    } else if targets.len() == 1 {
        "shared"
    } else {
        "distinct"
    };
    PointerObservation {
        same_target_pointer_consistent,
        distinct_target_pointer_isolated,
        label,
    }
}

fn snapshot_unique_target_bytes(snapshots: &[Snapshot]) -> usize {
    let mut targets = HashMap::<String, usize>::new();
    for anchor in snapshots.iter().flat_map(Snapshot::anchors) {
        targets.insert(
            anchor.target_part_name().as_str().to_owned(),
            anchor.target_bytes().len(),
        );
    }
    targets.values().copied().sum()
}

fn semantic_shape_ok(prepared: &Prepared, metrics: &GraphMetrics) -> bool {
    metrics.anchors == prepared.recipe.anchors
        && metrics.unique_targets == prepared.recipe.unique_targets
        && metrics.inbound_edges == prepared.recipe.edges
        && metrics.outbound_edges == prepared.recipe.outbound_edges
        && metrics.baseline_unique_target_bytes
            == prepared.recipe.target_bytes * prepared.recipe.unique_targets
        && metrics.unique_target_bytes == metrics.expected_unique_target_bytes
        && metrics.action_count > 0
        && metrics.action_group_count > 0
}

fn preservation_ok(prepared: &Prepared, metrics: &GraphMetrics) -> bool {
    let topology_ok = match prepared.recipe.topology.as_str() {
        "shared" | "case_equivalent_shared" => {
            metrics.same_target_pointer_consistent
                && metrics.distinct_target_pointer_isolated
                && metrics.shared_pointer_observation == "shared"
        },
        "distinct" => {
            metrics.same_target_pointer_consistent
                && metrics.distinct_target_pointer_isolated
                && metrics.shared_pointer_observation == "distinct"
        },
        _ => true,
    };
    let opaque_ok = !prepared.recipe.opaque_mce
        || (metrics.owner_xml_bytes > 0
            && metrics.retained_profile_bytes > 0
            && metrics.opaque_choice_preserved
            && metrics.opaque_fallback_preserved
            && metrics.opaque_payload_preserved
            && metrics.opaque_default_namespace_preserved
            && metrics.opaque_prefix_preserved
            && metrics.opaque_unknown_requires_preserved);
    topology_ok
        && opaque_ok
        && (!prepared.recipe.unknown_internal_outbound
            || metrics.unknown_internal_outbound_preserved)
        && (!prepared.recipe.unknown_external_outbound
            || metrics.unknown_external_outbound_preserved)
}

fn same_snapshot(left: &Snapshot, right: &Snapshot) -> bool {
    left.slide_index() == right.slide_index()
        && left.slide_part_name() == right.slide_part_name()
        && left.source_xml() == right.source_xml()
        && left.anchors().len() == right.anchors().len()
        && left
            .anchors()
            .iter()
            .zip(right.anchors())
            .all(|(left, right)| {
                left.target_part_name() == right.target_part_name()
                    && left.target_bytes() == right.target_bytes()
                    && left.profile().source() == right.profile().source()
                    && left.owner_xml() == right.owner_xml()
                    && left.outbound_references() == right.outbound_references()
            })
}

fn limits_json(limits: Limits) -> Value {
    json!({
        "anchors": limits.anchors,
        "target_bytes": limits.target_bytes,
        "total_target_bytes": limits.total_target_bytes,
        "target_relationships": limits.target_relationships,
    })
}

fn read_limits_json(limits: ReadLimits) -> Value {
    json!({
        "max_input_bytes": limits.max_input_bytes(),
        "max_archive_members": limits.max_archive_members(),
        "max_archive_total_entries": limits.max_archive_total_entries(),
        "max_archive_member_name_bytes": limits.max_archive_member_name_bytes(),
        "max_archive_metadata_bytes": limits.max_archive_metadata_bytes(),
        "max_archive_compressed_bytes": limits.max_archive_compressed_bytes(),
        "max_archive_entry_bytes": limits.max_archive_entry_bytes(),
        "max_archive_total_bytes": limits.max_archive_total_bytes(),
        "max_parts": limits.max_parts(),
        "max_part_bytes": limits.max_part_bytes(),
        "max_total_part_bytes": limits.max_total_part_bytes(),
        "max_content_types_bytes": limits.max_content_types_bytes(),
        "max_content_type_mappings": limits.max_content_type_mappings(),
        "max_relationship_parts": limits.max_relationship_parts(),
        "max_relationship_xml_bytes": limits.max_relationship_xml_bytes(),
        "max_total_relationship_xml_bytes": limits.max_total_relationship_xml_bytes(),
        "max_relationships_per_part": limits.max_relationships_per_part(),
        "max_total_relationships": limits.max_total_relationships(),
        "max_relationship_graph_nodes": limits.max_relationship_graph_nodes(),
        "max_xml_events": limits.max_xml_events(),
        "max_total_relationship_xml_events": limits.max_total_relationship_xml_events(),
        "max_xml_depth": limits.max_xml_depth(),
        "max_xml_attribute_bytes": limits.max_xml_attribute_bytes(),
        "max_relationship_target_bytes": limits.max_relationship_target_bytes(),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;

    #[test]
    fn retained_prepared_metadata_is_charged_before_baseline_snapshot() {
        let manifest = manifest();
        let lane = manifest
            .lanes
            .iter()
            .find(|lane| lane.id == "package_read_tiny_shared")
            .expect("representative lane is present");
        let recipe = manifest
            .recipes
            .iter()
            .find(|recipe| recipe.id == lane.recipe)
            .expect("representative recipe is present");
        let limits = owner_limits_for(recipe, lane).expect("representative limits are valid");
        let (prepared, _setup) =
            prepare_with_phase(lane, recipe, limits, ReadLimits::default()).expect("setup");

        assert_eq!(
            AllocSnapshot::now().live_bytes(),
            prepared.retained_baseline_live_bytes,
            "retained Prepared metadata must be included before the baseline snapshot",
        );
    }

    #[test]
    fn repeated_inbound_edges_do_not_make_distinct_targets_shared() {
        let targets = HashMap::from([
            (
                "/ppt/custom/action1.xml".to_owned(),
                TargetMetrics {
                    inbound_edges: 2,
                    pointer: Some(0x10),
                    pointer_consistent: true,
                    ..TargetMetrics::default()
                },
            ),
            (
                "/ppt/custom/action2.xml".to_owned(),
                TargetMetrics {
                    inbound_edges: 3,
                    pointer: Some(0x20),
                    pointer_consistent: true,
                    ..TargetMetrics::default()
                },
            ),
        ]);

        let observation = classify_pointer_observations(&targets);

        assert!(observation.same_target_pointer_consistent);
        assert!(observation.distinct_target_pointer_isolated);
        assert_eq!(observation.label, "distinct");
    }

    #[test]
    fn inconsistent_pointer_observations_are_not_accepted_as_topology() {
        let targets = HashMap::from([
            (
                "/ppt/custom/action1.xml".to_owned(),
                TargetMetrics {
                    pointer: Some(0x10),
                    pointer_consistent: false,
                    ..TargetMetrics::default()
                },
            ),
            (
                "/ppt/custom/action2.xml".to_owned(),
                TargetMetrics {
                    pointer: Some(0x10),
                    pointer_consistent: true,
                    ..TargetMetrics::default()
                },
            ),
        ]);

        let observation = classify_pointer_observations(&targets);

        assert!(!observation.same_target_pointer_consistent);
        assert!(!observation.distinct_target_pointer_isolated);
        assert_eq!(observation.label, "inconsistent");
    }

    #[test]
    fn inverse_uses_working_package_outside_retained_baseline() {
        let manifest = manifest();
        let lane = manifest
            .lanes
            .iter()
            .find(|lane| lane.id == "package_inverse_small_shared")
            .expect("inverse lane is present");
        let recipe = manifest
            .recipes
            .iter()
            .find(|recipe| recipe.id == lane.recipe)
            .expect("inverse recipe is present");
        let limits = owner_limits_for(recipe, lane).expect("inverse limits are valid");
        let (prepared, setup) =
            prepare_with_phase(lane, recipe, limits, ReadLimits::default()).expect("setup");
        assert!(prepared.operation_package.is_some());
        let receipt =
            run_sample(prepared, setup, false, lane.expected.as_deref()).expect("inverse sample");
        assert!(receipt.inverse_ok);
        assert!(receipt.baseline_reopenable);
        assert!(receipt.retained_baseline_balance_ok);
    }

    #[test]
    fn stale_content_type_checks_unmutated_source_baseline() {
        let manifest = manifest();
        let lane = manifest
            .lanes
            .iter()
            .find(|lane| lane.id == "stale_content_type")
            .expect("stale content-type lane is present");
        let recipe = manifest
            .recipes
            .iter()
            .find(|recipe| recipe.id == lane.recipe)
            .expect("stale content-type recipe is present");
        let limits = owner_limits_for(recipe, lane).expect("stale limits are valid");
        let (prepared, setup) =
            prepare_with_phase(lane, recipe, limits, ReadLimits::default()).expect("setup");
        let receipt = run_sample(prepared, setup, false, lane.expected.as_deref())
            .expect("stale content-type sample");
        assert_eq!(
            receipt.actual_error_type.as_deref(),
            Some("Error::ContentType")
        );
        assert!(receipt.source_unchanged_on_refusal);
        assert!(receipt.baseline_reopenable);
        assert!(receipt.retained_baseline_balance_ok);
    }

    #[test]
    fn refused_limit_lane_keeps_valid_source_baseline_reopenable() {
        let manifest = manifest();
        let lane = manifest
            .lanes
            .iter()
            .find(|lane| lane.id == "limit_anchor_one_under")
            .expect("negative limit lane is present");
        let recipe = manifest
            .recipes
            .iter()
            .find(|recipe| recipe.id == lane.recipe)
            .expect("negative limit recipe is present");
        let limits = owner_limits_for(recipe, lane).expect("negative limit is valid");
        let (prepared, setup) =
            prepare_with_phase(lane, recipe, limits, ReadLimits::default()).expect("setup");
        let receipt = run_sample(prepared, setup, false, lane.expected.as_deref())
            .expect("negative limit sample");
        assert_eq!(receipt.actual_error_type.as_deref(), Some("Error::Limit"));
        assert!(receipt.semantic_ok);
        assert!(receipt.baseline_reopenable);
        assert!(receipt.retained_baseline_balance_ok);
    }
}
