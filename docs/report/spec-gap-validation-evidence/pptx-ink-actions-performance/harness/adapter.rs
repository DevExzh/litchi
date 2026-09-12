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
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, ReadLimits, TargetMode};
use litchi_pptx::presentation::embedded::ink_actions::{Commit, Limits, Patch, Snapshot};
use litchi_pptx::{Error as PptxError, Package};
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};

use crate::support::{self, AllocSnapshot, PhaseRecord};

type BoxError = Box<dyn StdError + Send + Sync>;
type Result<T> = std::result::Result<T, BoxError>;

const SCHEMA: &str = "pptx-ink-actions-performance-v1";
const SOURCE_COMMIT: &str = "cf6fdb8e91dd232d7d762596763d2e9d8a5b9dbd";
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

const MANIFEST_BYTES: &[u8] = include_bytes!("../corpus-manifest.json");
const GENERATOR_SOURCE_BYTES: &[u8] = include_bytes!("adapter.rs");

#[derive(Clone, Debug, Deserialize)]
struct Manifest {
    schema: String,
    owner_commit: String,
    fixture_authority: FixtureAuthority,
    recipes: Vec<RecipeSpec>,
    lanes: Vec<LaneSpec>,
    measurement_gate: MeasurementGate,
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
    slide_names: Vec<PackURI>,
    target_names: Vec<PackURI>,
}

struct Prepared {
    lane: LaneSpec,
    recipe: RecipeSpec,
    package: Package,
    facts: FixtureFacts,
    package_before_bytes: Vec<u8>,
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
    debug: String,
    display: String,
    resource: Option<String>,
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
    expected_error: Option<String>,
    actual_error_type: Option<String>,
    actual_error_resource: Option<String>,
    actual_error_debug: Option<String>,
    actual_error_display: Option<String>,
}

#[derive(Clone, Debug)]
struct GraphMetrics {
    source_bytes: usize,
    owner_xml_bytes: usize,
    unique_target_bytes: usize,
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
    outbound_diagnostic_modes: Vec<String>,
    unknown_internal_outbound_preserved: bool,
    unknown_external_outbound_preserved: bool,
}

#[derive(Clone, Debug)]
struct ValidationReport {
    summary: ValidationSummary,
    metrics: GraphMetrics,
    output_bytes: Option<usize>,
}

#[derive(Clone, Debug, Serialize)]
struct SampleReceipt {
    semantic_ok: bool,
    preservation_ok: bool,
    inverse_ok: bool,
    source_unchanged_on_refusal: bool,
    expected_error: Option<String>,
    actual_error_type: Option<String>,
    actual_error_resource: Option<String>,
    actual_error_debug: Option<String>,
    actual_error_display: Option<String>,
    source_sha256: String,
    source_fnv1a64: u64,
    source_bytes: usize,
    owner_xml_bytes: usize,
    unique_target_bytes: usize,
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
    outbound_diagnostic_modes: Vec<String>,
    unknown_internal_outbound_preserved: bool,
    unknown_external_outbound_preserved: bool,
    phases: PhaseSet,
}

#[derive(Clone, Debug, Serialize)]
struct PhaseSet {
    setup: PhaseRecord,
    operation: PhaseRecord,
    validation: PhaseRecord,
    drop: PhaseRecord,
    postdrop: PhaseRecord,
}

#[derive(Clone, Debug, Serialize)]
struct Receipt {
    schema: &'static str,
    lane: String,
    recipe_id: String,
    source_commit: &'static str,
    helper_sha256: &'static str,
    generator_sha256: Option<String>,
    fixture_path: String,
    fixture_git_blob: String,
    generator_source_sha256: String,
    source_sha256: String,
    source_fnv1a64: u64,
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
        let _ = run_sample(prepared, setup, true)?;
    }

    let mut receipts = Vec::with_capacity(samples);
    for _ in 0..samples {
        let (prepared, setup) = prepare_with_phase(&lane, &recipe, owner_limits, read_limits)?;
        receipts.push(run_sample(prepared, setup, false)?);
    }

    let receipt = Receipt {
        schema: SCHEMA,
        lane: lane.id,
        recipe_id: recipe.id,
        source_commit: SOURCE_COMMIT,
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
    let mut successful = 0usize;
    let mut refusals = 0usize;
    for lane in &manifest.lanes {
        let recipe = manifest
            .recipes
            .iter()
            .find(|recipe| recipe.id == lane.recipe)
            .ok_or_else(|| format!("lane recipe is absent: {}", lane.recipe))?;
        let limits = owner_limits_for(recipe, lane)?;
        let (prepared, setup) = prepare_with_phase(lane, recipe, limits, ReadLimits::default())?;
        let receipt = run_sample(prepared, setup, false)?;
        if receipt.semantic_ok && receipt.preservation_ok {
            successful += 1;
        } else {
            refusals += 1;
        }
    }
    serde_json::to_string(&json!({
        "schema": "pptx-ink-actions-correctness-matrix-v1",
        "timings_collected": false,
        "source_commit": SOURCE_COMMIT,
        "helper_sha256": HELPER_SHA256,
        "recipe_count": manifest.recipes.len(),
        "lane_count": manifest.lanes.len(),
        "successful_lanes": successful,
        "refusal_or_failed_lanes": refusals,
    }))
    .map_err(Into::into)
}

pub fn host_probe() -> Result<String> {
    let manifest = manifest();
    validate_manifest(&manifest)?;
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
    serde_json::to_string(&json!({
        "schema": "pptx-ink-actions-host-probe-v1",
        "source_commit": SOURCE_COMMIT,
        "helper_sha256": HELPER_SHA256,
        "package_route": "Package::ink_actions",
        "presentation_route": "Presentation::ink_actions",
        "package_anchors": package_anchors,
        "presentation_anchors": presentation_anchors,
        "native_powerpoint_claim": false,
        "synthetic_complete_opc": true,
    }))
    .map_err(Into::into)
}

fn validate_manifest(manifest: &Manifest) -> Result<()> {
    if manifest.schema != "pptx-ink-actions-performance-scaffold-v2" {
        return Err(format!("unexpected manifest schema: {}", manifest.schema).into());
    }
    if manifest.owner_commit != SOURCE_COMMIT {
        return Err(format!("manifest owner pin differs: {}", manifest.owner_commit).into());
    }
    if manifest.fixture_authority.sha256 != HELPER_SHA256 {
        return Err("manifest helper hash differs from compiled provenance".into());
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
    } else if id == "opaque_mce_scalar_edit"
        || id == "unknown_outbound_read_edit"
        || id == "strict_shared_edit"
    {
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

    let target_opc = if recipe.stale_input.is_some() {
        let (mutated, _) = build_opc(recipe)?;
        let mut mutated = mutated;
        mutate_stale(&mut mutated, recipe.stale_input.as_deref().unwrap())?;
        mutated
    } else {
        build_opc(recipe)?.0
    };
    let mut package = Package::from_opc_package(target_opc)
        .map_err(|error| format!("public mutable package setup failed: {error:?}"))?;
    let package_before_bytes = package
        .to_bytes()
        .map_err(|error| format!("package baseline serialization failed: {error:?}"))?;

    Ok(Prepared {
        lane: lane.clone(),
        recipe: recipe.clone(),
        package,
        facts: baseline_facts,
        package_before_bytes,
        read_limits,
        owner_limits,
        source_snapshot,
        commit,
        inverse_patch,
        operation,
    })
}

fn prepare_with_phase(
    lane: &LaneSpec,
    recipe: &RecipeSpec,
    owner_limits: Limits,
    read_limits: ReadLimits,
) -> Result<(Prepared, PhaseRecord)> {
    let before = AllocSnapshot::now();
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
        let target = BlobPart::new(
            target_name.clone(),
            "text/xml".to_owned(),
            action_payload(recipe.target_bytes)?,
        );
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
            slide_xml(strict, &relationship_ids),
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
                    "opaque-dependency.xml".to_owned(),
                    "rIdOpaqueInternal".to_owned(),
                    false,
                );
            package.add_part(Box::new(BlobPart::new(
                PackURI::new("/ppt/custom/opaque-dependency.xml")
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
    Ok((
        package,
        FixtureFacts {
            source_bytes,
            source_sha256,
            source_fnv1a64,
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

fn action_payload(target_bytes: usize) -> Result<Vec<u8>> {
    let prefix = format!(
        r###"<?xml version="1.0" encoding="UTF-8"?><ia:actions xmlns:ia="{ACTION}" xmlns:i="{INKML}" xmlns:v="urn:vendor" lengthUnit="cm" timeUnit="ms" xml:id="root"><i:definitions><v:future v:flag="keep"/></i:definitions><ia:action xml:id="a0" type="add" startTime="0"><ia:property name="kind" value="old"/><ia:actionData xml:id="d0" name="stroke" ref="#d1"><ia:transform matrix="1,0,0,1"/><i:trace><v:opaque>"###
    );
    let suffix = r###"</v:opaque></i:trace><i:traceView/></ia:actionData><ia:actionDataGroup xml:id="dg0" name="group"><ia:actionData xml:id="d1" name="other"/></ia:actionDataGroup></ia:action><ia:actionGroup xml:id="ag0" type="transform" startTime="1"><ia:action xml:id="a1" type="remove" startTime="2"><ia:actionData/></ia:action></ia:actionGroup></ia:actions>"###;
    if target_bytes < prefix.len() + suffix.len() {
        return Err(format!("target recipe is below valid payload minimum: {target_bytes}").into());
    }
    let filler = "x".repeat(target_bytes - prefix.len() - suffix.len());
    let mut output = Vec::with_capacity(target_bytes);
    output.extend_from_slice(prefix.as_bytes());
    output.extend_from_slice(filler.as_bytes());
    output.extend_from_slice(suffix.as_bytes());
    debug_assert_eq!(output.len(), target_bytes);
    Ok(output)
}

fn action_owner_anchor(relationship_id: &str) -> String {
    format!(
        r#"<compat:AlternateContent><compat:Choice Requires="p14main inkAction"><pp:contentPart rr:id="{relationship_id}"/></compat:Choice><compat:Fallback><pp:pic><pp:nvPicPr><pp:cNvPr id="42" name="Fallback"/><pp:cNvPicPr/><pp:nvPr/></pp:nvPicPr><pp:blipFill/><pp:spPr/></pp:pic></compat:Fallback></compat:AlternateContent>"#
    )
}

fn slide_xml(strict: bool, relationship_ids: &[String]) -> Vec<u8> {
    let anchors = relationship_ids
        .iter()
        .map(|relationship_id| action_owner_anchor(relationship_id))
        .collect::<String>();
    let pml = if strict { STRICT_PML } else { PML };
    let rel = if strict { STRICT_REL } else { REL };
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><pp:sld xmlns:pp="{pml}" xmlns:rr="{rel}" xmlns:compat="{MCE}" xmlns:p14main="{P14_MAIN}" xmlns:inkAction="{ACTION}" compat:Ignorable="p14main inkAction"><pp:cSld><pp:spTree><pp:nvGrpSpPr/><pp:grpSpPr/>{anchors}</pp:spTree></pp:cSld></pp:sld>"#
    )
    .into_bytes()
}

fn serialize_opc(package: &OpcPackage) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    package
        .to_stream(&mut bytes)
        .map_err(|error| format!("OPC serialization failed: {error:?}"))?;
    Ok(bytes)
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
            bytes.push(b' ');
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
                bytes.push(b' ');
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

fn run_sample(mut prepared: Prepared, setup: PhaseRecord, _warmup: bool) -> Result<SampleReceipt> {
    let operation_before = AllocSnapshot::now();
    let operation_started = Instant::now();
    let outcome = execute_operation(&mut prepared)?;
    let operation_after = AllocSnapshot::now();
    let operation_phase = operation_before.delta(operation_after, operation_started.elapsed());

    let validation_before = AllocSnapshot::now();
    let validation_started = Instant::now();
    let report = validate_outcome(&mut prepared, &outcome)?;
    let validation_after = AllocSnapshot::now();
    let validation_phase = validation_before.delta(validation_after, validation_started.elapsed());

    let drop_before = AllocSnapshot::now();
    let drop_started = Instant::now();
    drop(outcome);
    let drop_after = AllocSnapshot::now();
    let drop_phase = drop_before.delta(drop_after, drop_started.elapsed());

    let postdrop_before = AllocSnapshot::now();
    let postdrop_started = Instant::now();
    let source_sha256 = prepared.facts.source_sha256.clone();
    let source_fnv1a64 = prepared.facts.source_fnv1a64;
    let receipt = SampleReceipt {
        semantic_ok: report.summary.semantic_ok,
        preservation_ok: report.summary.preservation_ok,
        inverse_ok: report.summary.inverse_ok,
        source_unchanged_on_refusal: report.summary.source_unchanged_on_refusal,
        expected_error: report.summary.expected_error.clone(),
        actual_error_type: report.summary.actual_error_type.clone(),
        actual_error_resource: report.summary.actual_error_resource.clone(),
        actual_error_debug: report.summary.actual_error_debug.clone(),
        actual_error_display: report.summary.actual_error_display.clone(),
        source_sha256,
        source_fnv1a64,
        source_bytes: report.metrics.source_bytes,
        owner_xml_bytes: report.metrics.owner_xml_bytes,
        unique_target_bytes: report.metrics.unique_target_bytes,
        output_bytes: report.output_bytes,
        anchors: report.metrics.anchors,
        unique_targets: report.metrics.unique_targets,
        inbound_edges: report.metrics.inbound_edges,
        outbound_edges: report.metrics.outbound_edges,
        action_count: report.metrics.action_count,
        action_group_count: report.metrics.action_group_count,
        retained_owner_xml_bytes: report.metrics.retained_owner_xml_bytes,
        retained_target_bytes: report.metrics.retained_target_bytes,
        retained_profile_bytes: report.metrics.retained_profile_bytes,
        shared_pointer_observation: report.metrics.shared_pointer_observation,
        outbound_diagnostic_modes: report.metrics.outbound_diagnostic_modes,
        unknown_internal_outbound_preserved: report.metrics.unknown_internal_outbound_preserved,
        unknown_external_outbound_preserved: report.metrics.unknown_external_outbound_preserved,
        phases: PhaseSet {
            setup,
            operation: operation_phase,
            validation: validation_phase,
            drop: drop_phase,
            postdrop: empty_phase(),
        },
    };
    drop(prepared);
    let postdrop_after = AllocSnapshot::now();
    let mut receipt = receipt;
    receipt.phases.postdrop = postdrop_before.delta(postdrop_after, postdrop_started.elapsed());
    Ok(receipt)
}

fn execute_operation(prepared: &mut Prepared) -> Result<HeldOutcome> {
    let outcome = match prepared.operation {
        OperationKind::PackageRead => public_snapshots(prepared.package.ink_actions()),
        OperationKind::PresentationRead => public_snapshots(
            prepared
                .package
                .presentation()
                .and_then(|presentation| presentation.ink_actions()),
        ),
        OperationKind::PackageReadWithLimits => public_snapshots(
            prepared
                .package
                .ink_actions_with_limits(prepared.owner_limits),
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
                .patch();
            public_snapshot(prepared.package.apply_ink_actions_patch(patch))
        },
        OperationKind::Inverse => {
            let patch = prepared
                .commit
                .as_ref()
                .ok_or("inverse operation has no prepared patch")?
                .patch();
            let inverse = prepared
                .inverse_patch
                .as_ref()
                .ok_or("inverse operation has no inverse patch")?;
            match prepared.package.apply_ink_actions_patch(patch) {
                Ok(_) => public_snapshot(prepared.package.apply_ink_actions_patch(inverse)),
                Err(error) => HeldOutcome::Error(actual_error(&error)),
            }
        },
        OperationKind::Save => match prepared.package.to_bytes() {
            Ok(bytes) => HeldOutcome::Success(HeldResult::Bytes(bytes)),
            Err(error) => HeldOutcome::Error(actual_error(&error)),
        },
        OperationKind::SaveReopen => match prepared.package.to_bytes() {
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
            prepared
                .package
                .ink_actions_with_limits(prepared.owner_limits),
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
    let resource = debug
        .split("resource:")
        .nth(1)
        .and_then(|value| value.split(',').next())
        .map(|value| value.trim().trim_matches('"').to_owned());
    ActualError {
        debug,
        display,
        resource,
    }
}

fn validate_outcome(prepared: &mut Prepared, outcome: &HeldOutcome) -> Result<ValidationReport> {
    let expected_error = prepared.lane.expected.clone();
    let mut summary = ValidationSummary {
        semantic_ok: false,
        preservation_ok: false,
        inverse_ok: true,
        source_unchanged_on_refusal: false,
        expected_error: expected_error.clone(),
        actual_error_type: None,
        actual_error_resource: None,
        actual_error_debug: None,
        actual_error_display: None,
    };

    let mut metrics = empty_metrics(&prepared.facts);
    let mut output_bytes = None;
    match outcome {
        HeldOutcome::Error(error) => {
            summary.actual_error_type =
                Some(error.debug.lines().next().unwrap_or_default().to_owned());
            summary.actual_error_resource = error.resource.clone();
            summary.actual_error_debug = Some(error.debug.clone());
            summary.actual_error_display = Some(error.display.clone());
            let expected = expected_error.as_deref().unwrap_or("success");
            summary.semantic_ok = expected != "success" && error_matches(expected, error);
            summary.source_unchanged_on_refusal = prepared
                .package
                .to_bytes()
                .map(|bytes| bytes == prepared.package_before_bytes)
                .unwrap_or(false);
            summary.preservation_ok = summary.source_unchanged_on_refusal;
        },
        HeldOutcome::Success(result) => {
            let snapshots = match result {
                HeldResult::Snapshots(snapshots) => snapshots.clone(),
                HeldResult::Snapshot(snapshot) => vec![snapshot.clone()],
                HeldResult::Commit(commit) => vec![commit.snapshot().clone()],
                HeldResult::SaveReopen { bytes, snapshots } => {
                    output_bytes = Some(bytes.len());
                    snapshots.clone()
                },
                HeldResult::Bytes(bytes) => {
                    output_bytes = Some(bytes.len());
                    match Package::from_vec_with_limits(bytes.clone(), prepared.read_limits)
                        .and_then(|package| package.ink_actions_with_limits(prepared.owner_limits))
                    {
                        Ok(snapshots) => snapshots,
                        Err(error) => {
                            summary.actual_error_type = Some(
                                format!("{error:?}")
                                    .lines()
                                    .next()
                                    .unwrap_or_default()
                                    .to_owned(),
                            );
                            summary.actual_error_debug = Some(format!("{error:?}"));
                            summary.actual_error_display = Some(error.to_string());
                            Vec::new()
                        },
                    }
                },
            };
            if !snapshots.is_empty() {
                metrics = graph_metrics(&prepared.facts, &snapshots);
                summary.semantic_ok = semantic_shape_ok(prepared, &metrics);
                summary.preservation_ok = preservation_ok(prepared, &metrics);
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
    })
}

fn error_matches(expected: &str, actual: &ActualError) -> bool {
    if expected == "success" {
        return false;
    }
    let expected_type = expected
        .strip_prefix("Error::")
        .unwrap_or(expected)
        .split_whitespace()
        .next()
        .unwrap_or(expected);
    let type_match = match expected_type {
        "Opc(OpcError::SignedSourceRequiresExplicitPolicy)" => {
            actual.debug.contains("SignedSourceRequiresExplicitPolicy")
        },
        "ContentType" => actual.debug.contains("ContentType"),
        "StaleSource" => actual.debug.contains("StaleSource"),
        "Limit" => actual.debug.contains("Limit"),
        other => actual.debug.contains(other),
    };
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
        .map(|value| actual.debug.contains(&format!("limit: {}", value.trim())))
        .unwrap_or(true);
    type_match && resource_match && limit_match
}

fn empty_metrics(facts: &FixtureFacts) -> GraphMetrics {
    GraphMetrics {
        source_bytes: facts.source_bytes.len(),
        owner_xml_bytes: 0,
        unique_target_bytes: 0,
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
        outbound_diagnostic_modes: Vec::new(),
        unknown_internal_outbound_preserved: false,
        unknown_external_outbound_preserved: false,
    }
}

fn graph_metrics(facts: &FixtureFacts, snapshots: &[Snapshot]) -> GraphMetrics {
    let anchors = snapshots
        .iter()
        .flat_map(Snapshot::anchors)
        .collect::<Vec<_>>();
    let mut targets = HashMap::<String, (usize, usize, usize, usize, Vec<TargetMode>)>::new();
    let mut pointer_by_target = HashMap::<String, *const u8>::new();
    let mut shared = false;
    let mut owner_xml_bytes = 0usize;
    for anchor in &anchors {
        owner_xml_bytes = owner_xml_bytes.saturating_add(anchor.owner_xml().len());
        let target = anchor.target_part_name().as_str().to_owned();
        let entry = targets.entry(target.clone()).or_insert((
            anchor.target_bytes().len(),
            anchor.inbound_references().len(),
            anchor.outbound_references().len(),
            anchor.profile().source().len(),
            Vec::new(),
        ));
        entry.0 = anchor.target_bytes().len();
        if let Some(previous) = pointer_by_target.insert(target, anchor.target_bytes().as_ptr()) {
            shared |= previous == anchor.target_bytes().as_ptr();
        }
        for reference in anchor.outbound_references() {
            entry.4.push(reference.target_mode());
        }
    }
    let mut inbound_edges = 0usize;
    let mut outbound_edges = 0usize;
    let mut unique_target_bytes = 0usize;
    let mut retained_profile_bytes = 0usize;
    let mut action_count = 0usize;
    let mut action_group_count = 0usize;
    let mut modes = BTreeSet::new();
    for (target_bytes, inbound, outbound, profile_bytes, target_modes) in targets.values() {
        unique_target_bytes = unique_target_bytes.saturating_add(*target_bytes);
        inbound_edges = inbound_edges.saturating_add(*inbound);
        outbound_edges = outbound_edges.saturating_add(*outbound);
        retained_profile_bytes = retained_profile_bytes.saturating_add(*profile_bytes);
        if let Some(anchor) = anchors.iter().find(|anchor| {
            anchor.target_bytes().len() == *target_bytes
                && anchor.profile().source().len() == *profile_bytes
        }) {
            action_count = action_count.saturating_add(anchor.profile().actions().count());
            action_group_count =
                action_group_count.saturating_add(anchor.profile().action_groups().count());
        }
        for mode in target_modes {
            modes.insert(match mode {
                TargetMode::Internal => "internal_unknown",
                TargetMode::External => "external_unknown",
            });
        }
    }
    let expected_pointer_observation = if shared { "shared" } else { "distinct" };
    GraphMetrics {
        source_bytes: facts.source_bytes.len(),
        owner_xml_bytes,
        unique_target_bytes,
        anchors: anchors.len(),
        unique_targets: targets.len(),
        inbound_edges,
        outbound_edges,
        action_count,
        action_group_count,
        retained_owner_xml_bytes: owner_xml_bytes,
        retained_target_bytes: unique_target_bytes,
        retained_profile_bytes,
        shared_pointer_observation: expected_pointer_observation.to_owned(),
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
    }
}

fn semantic_shape_ok(prepared: &Prepared, metrics: &GraphMetrics) -> bool {
    metrics.anchors == prepared.recipe.anchors
        && metrics.unique_targets == prepared.recipe.unique_targets
        && metrics.inbound_edges == prepared.recipe.edges
        && metrics.outbound_edges == prepared.recipe.outbound_edges
        && metrics.unique_target_bytes
            == prepared.recipe.target_bytes * prepared.recipe.unique_targets
        && metrics.action_count > 0
        && metrics.action_group_count > 0
}

fn preservation_ok(prepared: &Prepared, metrics: &GraphMetrics) -> bool {
    let topology_ok = match prepared.recipe.topology.as_str() {
        "shared" | "case_equivalent_shared" => metrics.shared_pointer_observation == "shared",
        "distinct" => metrics.shared_pointer_observation == "distinct",
        _ => true,
    };
    let opaque_ok = !prepared.recipe.opaque_mce
        || (metrics.owner_xml_bytes > 0 && metrics.retained_profile_bytes > 0);
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
