//! Bounded correctness smoke for the public DOCX `stylesWithEffects` owner API.
//!
//! The adapter deliberately uses only the public `litchi_docx` and OPC APIs.
//! Fixture construction is kept here so the smoke can prove the package and
//! graph invariants without depending on DOCX test helpers or implementation
//! internals.  It is a correctness receipt generator; it is not a sealed
//! performance benchmark.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::print_stdout,
    clippy::shadow_reuse,
    reason = "the opt-in evidence harness owns fixture recipes and JSON receipts"
)]

use std::collections::{BTreeMap, BTreeSet, VecDeque};
use std::error::Error as StdError;
use std::fmt::Write as FmtWrite;
use std::io::{Cursor, Read, Write};
use std::sync::Arc;
use std::time::Instant;

use litchi_docx::styles::Type;
use litchi_docx::styles::effects::{Owner, Resource};
use litchi_docx::{Error as DocxError, Package};
use litchi_opc::phys_pkg::PhysPkgReader;
use litchi_opc::{OpcError, OpcPackage, PackURI, ReadLimits, ReadResource};
use quick_xml::Reader;
use quick_xml::events::Event;
use sha2::{Digest as _, Sha256};
use zip::CompressionMethod;
use zip::ZipArchive;
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

pub type BoxError = Box<dyn StdError + Send + Sync>;
type Result<T> = std::result::Result<T, BoxError>;

const BUG: &[u8] = include_bytes!("../fixtures/Bug54849.docx");
const SIGNED: &[u8] = include_bytes!("../fixtures/ms-office-2010-signed.docx");
const COMPLEX: &[u8] = include_bytes!("../fixtures/ComplexNumberedLists.docx");
const GLOSSARY: &[u8] = include_bytes!("../fixtures/testGlossary.docx");

const EFFECTS_CONTENT_TYPE: &str = "application/vnd.ms-word.stylesWithEffects+xml";
const TRANSITIONAL_W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

pub const LANES: &[&str] = &[
    "native_capture_bug_main",
    "native_capture_bug_glossary",
    "native_capture_signed_main",
    "native_capture_signed_glossary_absent",
    "native_capture_complex_main",
    "native_capture_complex_glossary_absent",
    "native_capture_glossary_main",
    "native_capture_glossary_glossary",
    "source_noop_main",
    "source_noop_glossary",
    "projection_main",
    "projection_glossary",
    "replace_main",
    "replace_glossary",
    "remove_main",
    "remove_glossary",
    "add_main_absent",
    "add_glossary_missing",
    "inverse_replace_main",
    "inverse_remove_main",
    "stale_patch_main",
    "signed_noop",
    "signed_changed",
    "independent_main",
    "independent_glossary",
    "cap_parts",
    "cap_total_part_bytes",
    "cap_total_relationships",
    "cap_total_relationship_xml_events",
    "cap_total_relationship_xml_bytes",
    "cap_relationship_parts",
    "cap_relationship_graph_nodes",
    "malformed_duplicate_owner",
    "malformed_third_orphan",
    "malformed_external",
    "malformed_wrong_content_type",
    "malformed_outbound",
    "malformed_shared_inbound",
    "malformed_root",
    "malformed_namespace",
    "malformed_opaque_xml",
    "malformed_unbound_descendant",
    "malformed_invalid_qname",
    "malformed_raw_attribute",
    "malformed_raw_text",
    "malformed_control",
    "malformed_invalid_char_ref",
    "malformed_empty_prefix",
    "malformed_reserved_xml_uri",
    "malformed_xml_version",
    "malformed_xml_events",
    "malformed_xml_depth",
];

const EXPECTED_REFUSALS: &[&str] = &[
    "add_glossary_missing",
    "signed_changed",
    "stale_patch_main",
    "malformed_duplicate_owner",
    "malformed_third_orphan",
    "malformed_external",
    "malformed_wrong_content_type",
    "malformed_outbound",
    "malformed_shared_inbound",
    "malformed_root",
    "malformed_namespace",
    "malformed_opaque_xml",
    "malformed_unbound_descendant",
    "malformed_invalid_qname",
    "malformed_raw_attribute",
    "malformed_raw_text",
    "malformed_control",
    "malformed_invalid_char_ref",
    "malformed_empty_prefix",
    "malformed_reserved_xml_uri",
    "malformed_xml_version",
    "malformed_xml_events",
    "malformed_xml_depth",
];

#[derive(Clone, Copy, Debug, Default)]
pub struct PhaseTimes {
    pub capture_ns: u64,
    pub snapshot_ns: u64,
    pub stage_ns: u64,
    pub commit_ns: u64,
    pub publish_ns: u64,
    pub reopen_ns: u64,
    pub inverse_reopen_ns: u64,
    pub inverse_ns: u64,
    pub projection_ns: u64,
    pub opaque_ns: u64,
    pub graph_ns: u64,
    pub readback_ns: u64,
    pub validation_ns: u64,
}

#[derive(Clone, Debug)]
pub struct ErrorReceipt {
    pub class: String,
    pub variant: String,
    pub message: String,
    pub typed_match: bool,
    pub resource: Option<String>,
    pub actual: Option<u64>,
    pub maximum: Option<u64>,
}

#[derive(Clone, Copy, Debug, Default)]
pub struct PackageMetrics {
    pub parts: u64,
    pub total_part_bytes: u64,
    pub total_relationships: u64,
    pub relationship_parts: u64,
    pub relationship_graph_nodes: u64,
    pub relationship_xml_bytes: u64,
    pub relationship_xml_events: u64,
}

#[derive(Clone, Debug)]
pub struct RunResult {
    pub actual_success: bool,
    pub ingress_refusal: bool,
    pub no_output_ok: bool,
    pub semantic_ok: bool,
    pub opaque_ok: bool,
    pub exact_inverse_ok: bool,
    pub source_readback_physical_ok: Option<bool>,
    pub source_readback_metadata_ok: Option<bool>,
    pub output_bytes: u64,
    pub input_sha256: String,
    pub output_sha256: Option<String>,
    pub input_metrics: PackageMetrics,
    pub output_metrics: Option<PackageMetrics>,
    pub input_member_digest: String,
    pub output_member_digest: Option<String>,
    pub cap_exact_fit_ok: Option<bool>,
    pub cap_under_refused_ok: Option<bool>,
    pub cap_refusal: Option<ErrorReceipt>,
    pub cap_commit_refusal: Option<ErrorReceipt>,
    pub cap_source_metrics: Option<PackageMetrics>,
    pub cap_projected_metrics: Option<PackageMetrics>,
    pub cap_existing: Option<CapEvidence>,
    pub phases: PhaseTimes,
    pub error: Option<ErrorReceipt>,
}

/// Boundary evidence for an existing-owner replacement.  Some limits count
/// package topology and therefore cannot grow when a resource is replaced in
/// an already bound part; those rows are explicitly marked inapplicable.
#[derive(Clone, Debug)]
pub struct CapEvidence {
    pub applicable: bool,
    pub source_metrics: PackageMetrics,
    pub projected_metrics: PackageMetrics,
    pub exact_fit_ok: bool,
    pub exact_opaque_ok: bool,
    pub source_unrelated_member_digest: Option<String>,
    pub exact_unrelated_member_digest: Option<String>,
    pub under_refused_ok: bool,
    pub commit_stage_checked: bool,
    pub refusal: Option<ErrorReceipt>,
    pub commit_refusal: Option<ErrorReceipt>,
}

#[derive(Clone)]
pub struct Fixture {
    pub name: &'static str,
    pub package: Arc<[u8]>,
    pub native: bool,
    pub signed: bool,
    pub main_present: bool,
    pub glossary_present: bool,
    pub expected_package_sha256: Option<&'static str>,
}

#[derive(Clone, Copy)]
struct ExpectedMember {
    name: &'static str,
    length: usize,
    sha256: &'static str,
}

#[derive(Clone, Copy)]
struct NativeExpectation {
    name: &'static str,
    package_sha256: &'static str,
    main: Option<ExpectedMember>,
    glossary: Option<ExpectedMember>,
}

const NATIVE_EXPECTATIONS: &[NativeExpectation] = &[
    NativeExpectation {
        name: "Bug54849.docx",
        package_sha256: "f54182713ea5ce5d77b9593d3d9d24e645460043cec0b40ef59c932385f084d3",
        main: Some(ExpectedMember {
            name: "word/stylesWithEffects.xml",
            length: 19_883,
            sha256: "799de1f7a4ce43f0ca101dc750e8a8bd6e75bcb721f7744d4787dad576cda3b1",
        }),
        glossary: Some(ExpectedMember {
            name: "word/glossary/stylesWithEffects.xml",
            length: 16_138,
            sha256: "d27f6ced340ffa173b3b861b08a4006e76673a411687b4145f5e46dc6dae13eb",
        }),
    },
    NativeExpectation {
        name: "ms-office-2010-signed.docx",
        package_sha256: "bc55c0362722818823a6dd95f8e0ca9869e179ace972a0915241feb4677bde5f",
        main: Some(ExpectedMember {
            name: "word/stylesWithEffects.xml",
            length: 15_710,
            sha256: "00c5cda7671bf545a8c97312f14b2b8bc0ee7fa469b36c25ee158c8a5c1c1568",
        }),
        glossary: None,
    },
    NativeExpectation {
        name: "ComplexNumberedLists.docx",
        package_sha256: "297a085a7d433af2eeee7661e8db21539452cb585096484774a1e9f5f258b0b6",
        main: Some(ExpectedMember {
            name: "word/stylesWithEffects.xml",
            length: 15_955,
            sha256: "b4bf5d355a45daf0a1085e73fe27041b5db22bfa23820f18ffaa7f9c8cb70f18",
        }),
        glossary: None,
    },
    NativeExpectation {
        name: "testGlossary.docx",
        package_sha256: "8ccd581d8f0ae102b220228ad26b3974821a7ce8e3ff7df4b78f7da8a0d06ed9",
        main: Some(ExpectedMember {
            name: "word/stylesWithEffects.xml",
            length: 20_117,
            sha256: "e72df38e71a351ebaaf7b102e8ab7862b7e5eb04cc86e187f8510746b70d3f54",
        }),
        glossary: Some(ExpectedMember {
            name: "word/glossary/stylesWithEffects.xml",
            length: 16_244,
            sha256: "78112f02ff4e0b94a6c99f8688d86c6fa91d54d6c57c504001453b9c73e3dad1",
        }),
    },
];

#[must_use]
pub fn is_known_lane(lane: &str) -> bool {
    LANES.contains(&lane)
}

#[must_use]
pub fn expected_success(lane: &str) -> bool {
    is_known_lane(lane) && !EXPECTED_REFUSALS.contains(&lane)
}

pub fn fixture_for_lane(lane: &str) -> Result<Fixture> {
    let (name, bytes, native, signed, main_present, glossary_present) = if lane.contains("bug")
        || matches!(
            lane,
            "source_noop_main"
                | "projection_main"
                | "replace_main"
                | "remove_main"
                | "inverse_replace_main"
                | "inverse_remove_main"
                | "stale_patch_main"
                | "independent_main"
                | "malformed_duplicate_owner"
                | "malformed_third_orphan"
                | "malformed_external"
                | "malformed_wrong_content_type"
                | "malformed_outbound"
                | "malformed_shared_inbound"
                | "malformed_root"
                | "malformed_namespace"
                | "malformed_opaque_xml"
                | "malformed_unbound_descendant"
                | "malformed_invalid_qname"
                | "malformed_raw_attribute"
                | "malformed_raw_text"
                | "malformed_control"
                | "malformed_invalid_char_ref"
                | "malformed_empty_prefix"
                | "malformed_reserved_xml_uri"
                | "malformed_xml_version"
                | "malformed_xml_events"
                | "malformed_xml_depth"
        ) {
        ("Bug54849.docx", BUG, true, false, true, true)
    } else if lane.contains("signed") {
        (
            "ms-office-2010-signed.docx",
            SIGNED,
            true,
            true,
            true,
            false,
        )
    } else if lane.contains("complex") || lane == "add_glossary_missing" {
        (
            "ComplexNumberedLists.docx",
            COMPLEX,
            true,
            false,
            true,
            false,
        )
    } else if lane.contains("glossary") {
        ("testGlossary.docx", GLOSSARY, true, false, true, true)
    } else if lane.starts_with("cap_") || lane == "add_main_absent" {
        let source = remove_main_effects(COMPLEX)?;
        let source = if lane.starts_with("cap_") {
            remove_member(&source, "word/_rels/document.xml.rels")?
        } else {
            source
        };
        return Ok(Fixture {
            name: "ComplexNumberedLists.docx:main-effects-absent",
            package: Arc::from(source),
            native: false,
            signed: false,
            main_present: false,
            glossary_present: false,
            expected_package_sha256: None,
        });
    } else {
        return Err(format!("unknown stylesWithEffects lane: {lane}").into());
    };

    Ok(Fixture {
        name,
        package: Arc::from(bytes),
        native,
        signed,
        main_present,
        glossary_present,
        expected_package_sha256: native_expectation(name).map(|value| value.package_sha256),
    })
}

fn native_expectation(name: &str) -> Option<NativeExpectation> {
    NATIVE_EXPECTATIONS
        .iter()
        .copied()
        .find(|expectation| expectation.name == name)
}

pub fn run_once(lane: &str, fixture: &Fixture) -> Result<RunResult> {
    if !is_known_lane(lane) {
        return Err(format!("unknown lane: {lane}").into());
    }
    match lane {
        name if name.starts_with("native_capture_") => run_native_capture(lane, fixture),
        "source_noop_main" => run_noop(fixture, Owner::MainDocument),
        "source_noop_glossary" => run_noop(fixture, Owner::Glossary),
        "projection_main" => run_projection(fixture, Owner::MainDocument),
        "projection_glossary" => run_projection(fixture, Owner::Glossary),
        "replace_main" => run_replace(fixture, Owner::MainDocument, false),
        "replace_glossary" => run_replace(fixture, Owner::Glossary, false),
        "remove_main" => run_remove(fixture, Owner::MainDocument, false),
        "remove_glossary" => run_remove(fixture, Owner::Glossary, false),
        "add_main_absent" => run_add_absent(fixture),
        "add_glossary_missing" => run_missing_glossary(fixture),
        "inverse_replace_main" => run_replace(fixture, Owner::MainDocument, true),
        "inverse_remove_main" => run_remove(fixture, Owner::MainDocument, true),
        "stale_patch_main" => run_stale_patch(fixture),
        "signed_noop" => run_signed_noop(fixture),
        "signed_changed" => run_signed_changed(fixture),
        "independent_main" => run_independent(fixture, Owner::MainDocument),
        "independent_glossary" => run_independent(fixture, Owner::Glossary),
        name if name.starts_with("cap_") => run_cap(fixture, cap_from_lane(name)?),
        "malformed_opaque_xml" => run_malformed_opaque_xml(fixture),
        name if matches!(
            name,
            "malformed_unbound_descendant"
                | "malformed_invalid_qname"
                | "malformed_raw_attribute"
                | "malformed_raw_text"
                | "malformed_control"
                | "malformed_invalid_char_ref"
                | "malformed_empty_prefix"
                | "malformed_reserved_xml_uri"
                | "malformed_xml_version"
        ) =>
        {
            run_malformed_xml_case(fixture, name)
        },
        "malformed_xml_events" => run_xml_limit_refusal(fixture, false),
        "malformed_xml_depth" => run_xml_limit_refusal(fixture, true),
        name if name.starts_with("malformed_") => run_malformed(fixture, name),
        _ => Err(format!("lane dispatch missing: {lane}").into()),
    }
}

fn run_native_capture(lane: &str, fixture: &Fixture) -> Result<RunResult> {
    let owner = if lane.ends_with("_glossary") || lane.ends_with("_glossary_absent") {
        Owner::Glossary
    } else {
        Owner::MainDocument
    };
    let capture = Instant::now();
    let package = Package::from_reader(Cursor::new(fixture.package.as_ref()))?;
    let snapshot = package.styles_with_effects(owner)?;
    let capture_ns = elapsed(capture);
    let validation = Instant::now();
    let expected_present = match owner {
        Owner::MainDocument => fixture.main_present,
        Owner::Glossary => fixture.glossary_present,
    };
    let mut semantic_ok = snapshot.is_empty() != expected_present;
    if let Some(resource) = snapshot.resource() {
        semantic_ok &= resource.conformance() == snapshot.conformance();
        semantic_ok &= !resource.xml_bytes().is_empty();
        semantic_ok &= resource.styles().len() == resource.projection().len();
    }
    if fixture.native {
        check_native_expectation(fixture)?;
        if let Some(expectation) = native_expectation(fixture.name) {
            let selected = match owner {
                Owner::MainDocument => expectation.main,
                Owner::Glossary => expectation.glossary,
            };
            if let Some(selected) = selected {
                let members = member_bytes(fixture.package.as_ref())?;
                let xml = members
                    .get(selected.name)
                    .ok_or_else(|| format!("missing native member {}", selected.name))?;
                semantic_ok &= xml.len() == selected.length;
                semantic_ok &= sha256_hex(xml) == selected.sha256;
                let ordinary = if owner == Owner::MainDocument {
                    "word/styles.xml"
                } else {
                    "word/glossary/styles.xml"
                };
                let ordinary_xml = members
                    .get(ordinary)
                    .ok_or_else(|| format!("missing ordinary styles member {ordinary}"))?;
                semantic_ok &= xml != ordinary_xml;
            }
        }
    }
    let validation_ns = elapsed(validation);
    let graph = Instant::now();
    let metrics = package_metrics(fixture.package.as_ref())?;
    let member_hashes = member_hashes(fixture.package.as_ref())?;
    let graph_ns = elapsed(graph);
    Ok(RunResult {
        actual_success: true,
        ingress_refusal: false,
        no_output_ok: true,
        semantic_ok,
        opaque_ok: semantic_ok,
        exact_inverse_ok: true,
        source_readback_physical_ok: None,
        source_readback_metadata_ok: None,
        output_bytes: fixture.package.len() as u64,
        input_sha256: sha256_hex(fixture.package.as_ref()),
        output_sha256: Some(sha256_hex(fixture.package.as_ref())),
        input_metrics: metrics,
        output_metrics: Some(metrics),
        input_member_digest: member_digest(&member_hashes),
        output_member_digest: Some(member_digest(&member_hashes)),
        cap_exact_fit_ok: None,
        cap_under_refused_ok: None,
        cap_refusal: None,
        cap_commit_refusal: None,
        cap_source_metrics: None,
        cap_projected_metrics: None,
        cap_existing: None,
        phases: PhaseTimes {
            capture_ns,
            graph_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
        error: None,
    })
}

fn run_noop(fixture: &Fixture, owner: Owner) -> Result<RunResult> {
    let mut package = open_package(fixture)?;
    let baseline = package_bytes(&mut package)?;
    let input_members = member_hashes(&baseline)?;
    let capture = Instant::now();
    let snapshot = package.styles_with_effects(owner)?;
    let capture_ns = elapsed(capture);
    let snapshot_phase = Instant::now();
    let commit = snapshot.edit().commit()?;
    let patch = commit.patch().clone();
    let snapshot_ns = elapsed(snapshot_phase);
    let publish = Instant::now();
    let applied = package.apply_styles_with_effects_patch(owner, &patch)?;
    let no_op = applied
        .resource()
        .cloned()
        .map(|resource| package.put_styles_with_effects(owner, resource))
        .transpose()?
        .is_none_or(|changed| !changed);
    let publish_ns = elapsed(publish);
    let output = package_bytes(&mut package)?;
    let validation = Instant::now();
    let output_members = member_hashes(&output)?;
    let semantic_ok = !commit.changed()
        && patch.is_empty()
        && no_op
        && output == baseline
        && snapshot_matches_package(&output, owner, snapshot.resource())?;
    let opaque_ok = output_members == input_members;
    let validation_ns = elapsed(validation);
    let graph = Instant::now();
    let metrics = package_metrics(&baseline)?;
    let graph_ns = elapsed(graph);
    Ok(success_result(
        fixture,
        baseline,
        output,
        input_members,
        output_members,
        metrics,
        metrics,
        semantic_ok,
        opaque_ok,
        true,
        PhaseTimes {
            capture_ns,
            snapshot_ns,
            publish_ns,
            graph_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
    ))
}

fn run_projection(fixture: &Fixture, owner: Owner) -> Result<RunResult> {
    let mut package = open_package(fixture)?;
    let baseline = package_bytes(&mut package)?;
    let input_members = member_hashes(&baseline)?;
    let capture = Instant::now();
    let snapshot = package.styles_with_effects(owner)?;
    let capture_ns = elapsed(capture);
    let projection_phase = Instant::now();
    let resource = snapshot
        .resource()
        .ok_or_else(|| format!("projection owner is absent: {owner}"))?;
    let projection = resource.projection();
    let style_count = projection.len();
    let first_id = projection
        .iter()
        .next()
        .map(|style| style.style_id().to_owned());
    let lookup_ok = first_id.as_deref().is_some_and(|id| {
        projection.get_by_id(id).is_some() && projection.resolved_numbering(id).is_ok()
    });
    let default_ok = projection.get_default(Type::Paragraph).is_none()
        || projection
            .get_default(Type::Paragraph)
            .is_some_and(|style| style.style_type() == Type::Paragraph);
    let projection_ns = elapsed(projection_phase);
    let validation = Instant::now();
    let semantic_ok = style_count > 0 && lookup_ok && default_ok;
    let validation_ns = elapsed(validation);
    let graph = Instant::now();
    let metrics = package_metrics(&baseline)?;
    let graph_ns = elapsed(graph);
    Ok(success_result(
        fixture,
        baseline.clone(),
        baseline,
        input_members.clone(),
        input_members,
        metrics,
        metrics,
        semantic_ok,
        true,
        true,
        PhaseTimes {
            capture_ns,
            projection_ns,
            graph_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
    ))
}

fn run_replace(fixture: &Fixture, owner: Owner, inverse: bool) -> Result<RunResult> {
    let mut package = open_package(fixture)?;
    let baseline = package_bytes(&mut package)?;
    let input_members = member_hashes(&baseline)?;
    let capture = Instant::now();
    let snapshot = package.styles_with_effects(owner)?;
    let capture_ns = elapsed(capture);
    let resource = snapshot
        .resource()
        .ok_or_else(|| format!("replace owner is absent: {owner}"))?;
    let replacement = changed_resource(resource)?;
    let stage = Instant::now();
    let mut edit = snapshot.edit();
    edit.replace_resource(Some(replacement.clone()))?;
    let stage_ns = elapsed(stage);
    let commit_start = Instant::now();
    let commit = edit.commit()?;
    let commit_ns = elapsed(commit_start);
    let patch = commit.patch().clone();
    let publish = Instant::now();
    package.apply_styles_with_effects_patch(owner, &patch)?;
    let publish_ns = elapsed(publish);
    let changed = package_bytes(&mut package)?;
    let reopen_start = Instant::now();
    let reopened = Package::from_reader(Cursor::new(changed.as_slice()))?;
    let reopened_snapshot = reopened.styles_with_effects(owner)?;
    let reopen_ns = elapsed(reopen_start);
    let validation = Instant::now();
    let semantic_ok = reopened_snapshot
        .resource()
        .is_some_and(|value| value.xml_bytes() == replacement.xml_bytes())
        && changed != baseline;
    let validation_ns = elapsed(validation);
    let opaque_phase = Instant::now();
    let opaque_ok = unchanged_except(
        &input_members,
        &member_hashes(&changed)?,
        &[effects_member(owner)],
    );
    let opaque_ns = elapsed(opaque_phase);
    let mut phases = PhaseTimes {
        capture_ns,
        stage_ns,
        commit_ns,
        publish_ns,
        reopen_ns,
        opaque_ns,
        validation_ns,
        ..PhaseTimes::default()
    };
    let (final_bytes, exact_inverse_ok, inverse_reopen_ns, inverse_ns) = if inverse {
        let inverse_start = Instant::now();
        let inverse_patch = patch.inverse();
        package.apply_styles_with_effects_patch(owner, &inverse_patch)?;
        let inverse_ns = elapsed(inverse_start);
        let restored = package_bytes(&mut package)?;
        let inverse_reopen_start = Instant::now();
        let restored_package = Package::from_reader(Cursor::new(restored.as_slice()))?;
        let restored_snapshot = restored_package.styles_with_effects(owner)?;
        let inverse_reopen_ns = elapsed(inverse_reopen_start);
        let exact = restored == baseline
            && restored_snapshot.resource().is_some_and(|value| {
                snapshot
                    .resource()
                    .is_some_and(|base| value.xml_bytes() == base.xml_bytes())
            });
        (restored, exact, inverse_reopen_ns, inverse_ns)
    } else {
        (changed, true, 0, 0)
    };
    phases.inverse_reopen_ns = inverse_reopen_ns;
    phases.inverse_ns = inverse_ns;
    let graph = Instant::now();
    let input_metrics = package_metrics(&baseline)?;
    let output_members = member_hashes(&final_bytes)?;
    let output_metrics = package_metrics(&final_bytes)?;
    let graph_ns = elapsed(graph);
    phases.graph_ns = graph_ns;
    let mut result = success_result(
        fixture,
        baseline,
        final_bytes,
        input_members,
        output_members,
        output_metrics,
        input_metrics,
        semantic_ok,
        opaque_ok,
        exact_inverse_ok,
        phases,
    );
    if inverse {
        result.opaque_ok &= result.output_sha256 == Some(result.input_sha256.clone());
    }
    Ok(result)
}

fn run_remove(fixture: &Fixture, owner: Owner, inverse: bool) -> Result<RunResult> {
    let mut package = open_package(fixture)?;
    let baseline = package_bytes(&mut package)?;
    let input_members = member_hashes(&baseline)?;
    let capture = Instant::now();
    let snapshot = package.styles_with_effects(owner)?;
    let capture_ns = elapsed(capture);
    let stage = Instant::now();
    let mut edit = snapshot.edit();
    edit.clear_resource();
    let stage_ns = elapsed(stage);
    let commit_start = Instant::now();
    let commit = edit.commit()?;
    let commit_ns = elapsed(commit_start);
    let patch = commit.patch().clone();
    let publish = Instant::now();
    package.apply_styles_with_effects_patch(owner, &patch)?;
    let publish_ns = elapsed(publish);
    let removed = package_bytes(&mut package)?;
    let reopen_start = Instant::now();
    let reopened = Package::from_reader(Cursor::new(removed.as_slice()))?;
    let removed_snapshot = reopened.styles_with_effects(owner)?;
    let reopen_ns = elapsed(reopen_start);
    let validation = Instant::now();
    let semantic_ok = removed_snapshot.is_empty()
        && !member_hashes(&removed)?.contains_key(effects_member(owner));
    let validation_ns = elapsed(validation);
    let opaque_phase = Instant::now();
    let opaque_ok = unchanged_except(
        &input_members,
        &member_hashes(&removed)?,
        &[
            effects_member(owner),
            owner_relationship_member(owner),
            target_relationship_member(owner),
            "[Content_Types].xml",
        ],
    );
    let opaque_ns = elapsed(opaque_phase);
    let mut phases = PhaseTimes {
        capture_ns,
        stage_ns,
        commit_ns,
        publish_ns,
        reopen_ns,
        opaque_ns,
        validation_ns,
        ..PhaseTimes::default()
    };
    let (final_bytes, exact_inverse_ok, inverse_reopen_ns, inverse_ns) = if inverse {
        let inverse_start = Instant::now();
        package.apply_styles_with_effects_patch(owner, &patch.inverse())?;
        let inverse_ns = elapsed(inverse_start);
        let restored = package_bytes(&mut package)?;
        let inverse_reopen_start = Instant::now();
        let restored_package = Package::from_reader(Cursor::new(restored.as_slice()))?;
        let restored_snapshot = restored_package.styles_with_effects(owner)?;
        let inverse_reopen_ns = elapsed(inverse_reopen_start);
        let exact = restored == baseline
            && restored_snapshot.resource().is_some_and(|value| {
                snapshot
                    .resource()
                    .is_some_and(|base| value.xml_bytes() == base.xml_bytes())
            });
        (restored, exact, inverse_reopen_ns, inverse_ns)
    } else {
        (removed, true, 0, 0)
    };
    phases.inverse_reopen_ns = inverse_reopen_ns;
    phases.inverse_ns = inverse_ns;
    let graph = Instant::now();
    let input_metrics = package_metrics(&baseline)?;
    let output_members = member_hashes(&final_bytes)?;
    let output_metrics = package_metrics(&final_bytes)?;
    let graph_ns = elapsed(graph);
    phases.graph_ns = graph_ns;
    let mut result = success_result(
        fixture,
        baseline,
        final_bytes,
        input_members,
        output_members,
        output_metrics,
        input_metrics,
        semantic_ok,
        opaque_ok,
        exact_inverse_ok,
        phases,
    );
    if inverse {
        result.opaque_ok &= result.output_sha256 == Some(result.input_sha256.clone());
    }
    Ok(result)
}

fn run_add_absent(fixture: &Fixture) -> Result<RunResult> {
    let mut package = open_package(fixture)?;
    let baseline = package_bytes(&mut package)?;
    let input_members = member_hashes(&baseline)?;
    let resource = fixture_resource(COMPLEX, "word/stylesWithEffects.xml")?;
    let capture = Instant::now();
    let snapshot = package.styles_with_effects(Owner::MainDocument)?;
    let capture_ns = elapsed(capture);
    if !snapshot.is_empty() {
        return Err("addition source unexpectedly has a main effects owner".into());
    }
    let stage = Instant::now();
    let mut edit = snapshot.edit();
    edit.replace_resource(Some(resource.clone()))?;
    let stage_ns = elapsed(stage);
    let commit_start = Instant::now();
    let commit = edit.commit()?;
    let commit_ns = elapsed(commit_start);
    let publish = Instant::now();
    package.apply_styles_with_effects_patch(Owner::MainDocument, commit.patch())?;
    let publish_ns = elapsed(publish);
    let output = package_bytes(&mut package)?;
    let reopen_start = Instant::now();
    let reopened = Package::from_reader(Cursor::new(output.as_slice()))?;
    let snapshot_after = reopened.styles_with_effects(Owner::MainDocument)?;
    let reopen_ns = elapsed(reopen_start);
    let validation = Instant::now();
    let semantic_ok = snapshot_after
        .resource()
        .is_some_and(|value| value.xml_bytes() == resource.xml_bytes())
        && member_hashes(&output)?.contains_key(effects_member(Owner::MainDocument));
    let validation_ns = elapsed(validation);
    let opaque_phase = Instant::now();
    let opaque_ok = unchanged_except(
        &input_members,
        &member_hashes(&output)?,
        &[
            effects_member(Owner::MainDocument),
            owner_relationship_member(Owner::MainDocument),
            target_relationship_member(Owner::MainDocument),
            "[Content_Types].xml",
        ],
    );
    let opaque_ns = elapsed(opaque_phase);
    let graph = Instant::now();
    let input_metrics = package_metrics(&baseline)?;
    let output_members = member_hashes(&output)?;
    let output_metrics = package_metrics(&output)?;
    let graph_ns = elapsed(graph);
    Ok(success_result(
        fixture,
        baseline,
        output.clone(),
        input_members,
        output_members,
        output_metrics,
        input_metrics,
        semantic_ok,
        opaque_ok,
        true,
        PhaseTimes {
            capture_ns,
            stage_ns,
            commit_ns,
            publish_ns,
            reopen_ns,
            opaque_ns,
            graph_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
    ))
}

fn run_missing_glossary(fixture: &Fixture) -> Result<RunResult> {
    let mut package = open_package(fixture)?;
    let baseline = package_bytes(&mut package)?;
    let input_members = member_hashes(&baseline)?;
    let resource = fixture_resource(SIGNED, "word/stylesWithEffects.xml")?;
    let capture = Instant::now();
    let snapshot = package.styles_with_effects(Owner::Glossary)?;
    let capture_ns = elapsed(capture);
    let stage = Instant::now();
    let error = package
        .put_styles_with_effects(Owner::Glossary, resource)
        .expect_err("missing glossary publication unexpectedly succeeded");
    let stage_ns = elapsed(stage);
    let validation = Instant::now();
    let error_receipt = classify_error(&error, ExpectedError::GlossaryMissing);
    let validation_ns = elapsed(validation);
    let (physical, metadata, readback_ns) = readback_unchanged(
        &mut package,
        &baseline,
        Owner::Glossary,
        snapshot.resource(),
    )?;
    let graph = Instant::now();
    let metrics = package_metrics(&baseline)?;
    let graph_ns = elapsed(graph);
    Ok(RunResult {
        actual_success: false,
        ingress_refusal: false,
        no_output_ok: true,
        semantic_ok: error_receipt.typed_match && physical && metadata,
        opaque_ok: metadata,
        exact_inverse_ok: false,
        source_readback_physical_ok: Some(physical),
        source_readback_metadata_ok: Some(metadata),
        output_bytes: 0,
        input_sha256: sha256_hex(&baseline),
        output_sha256: None,
        input_metrics: metrics,
        output_metrics: None,
        input_member_digest: member_digest(&input_members),
        output_member_digest: None,
        cap_exact_fit_ok: None,
        cap_under_refused_ok: None,
        cap_refusal: None,
        cap_commit_refusal: None,
        cap_source_metrics: None,
        cap_projected_metrics: None,
        cap_existing: None,
        phases: PhaseTimes {
            capture_ns,
            stage_ns,
            graph_ns,
            validation_ns,
            readback_ns,
            ..PhaseTimes::default()
        },
        error: Some(error_receipt),
    })
}

fn run_stale_patch(fixture: &Fixture) -> Result<RunResult> {
    let mut package = open_package(fixture)?;
    let baseline = package_bytes(&mut package)?;
    let input_members = member_hashes(&baseline)?;
    let snapshot = package.styles_with_effects(Owner::MainDocument)?;
    let replacement = changed_resource(snapshot.resource().ok_or("stale source missing")?)?;
    let mut edit = snapshot.edit();
    edit.replace_resource(Some(replacement))?;
    let patch = edit.commit()?.into_patch();
    package.apply_styles_with_effects_patch(Owner::MainDocument, &patch)?;
    let changed = package_bytes(&mut package)?;
    let changed_members = member_hashes(&changed)?;
    let refusal = Instant::now();
    let error = package
        .apply_styles_with_effects_patch(Owner::MainDocument, &patch)
        .expect_err("stale patch unexpectedly succeeded");
    let stage_ns = elapsed(refusal);
    let validation = Instant::now();
    let error_receipt = classify_error(&error, ExpectedError::StalePatch);
    let validation_ns = elapsed(validation);
    let readback = Instant::now();
    let after = package_bytes(&mut package)?;
    let physical = after == changed;
    let reopened = Package::from_reader(Cursor::new(after.as_slice()))?;
    let metadata = reopened
        .styles_with_effects(Owner::MainDocument)?
        .resource()
        .is_some();
    let readback_ns = elapsed(readback);
    let graph = Instant::now();
    let input_metrics = package_metrics(&baseline)?;
    let graph_ns = elapsed(graph);
    Ok(RunResult {
        actual_success: false,
        ingress_refusal: false,
        no_output_ok: true,
        semantic_ok: error_receipt.typed_match && physical && metadata,
        opaque_ok: changed_members == member_hashes(&after)?,
        exact_inverse_ok: false,
        source_readback_physical_ok: Some(physical),
        source_readback_metadata_ok: Some(metadata),
        output_bytes: 0,
        input_sha256: sha256_hex(&baseline),
        output_sha256: None,
        input_metrics,
        output_metrics: None,
        input_member_digest: member_digest(&input_members),
        output_member_digest: None,
        cap_exact_fit_ok: None,
        cap_under_refused_ok: None,
        cap_refusal: None,
        cap_commit_refusal: None,
        cap_source_metrics: None,
        cap_projected_metrics: None,
        cap_existing: None,
        phases: PhaseTimes {
            stage_ns,
            validation_ns,
            graph_ns,
            readback_ns,
            ..PhaseTimes::default()
        },
        error: Some(error_receipt),
    })
}

fn run_signed_noop(fixture: &Fixture) -> Result<RunResult> {
    let capture = Instant::now();
    let mut package = open_package(fixture)?;
    if !package.is_signed() {
        return Err("signed fixture did not report a signature".into());
    }
    let baseline = package_bytes(&mut package)?;
    let input_members = member_hashes(&baseline)?;
    let snapshot = package.styles_with_effects(Owner::MainDocument)?;
    let resource = snapshot
        .resource()
        .ok_or("signed main effects missing")?
        .clone();
    let capture_ns = elapsed(capture);
    let publish = Instant::now();
    let changed = package.put_styles_with_effects(Owner::MainDocument, resource)?;
    let publish_ns = elapsed(publish);
    let output = package_bytes(&mut package)?;
    let validation = Instant::now();
    let semantic_ok = !changed && package.is_signed() && output == baseline;
    let validation_ns = elapsed(validation);
    let opaque_phase = Instant::now();
    let output_members = member_hashes(&output)?;
    let opaque_ok = output_members == input_members;
    let opaque_ns = elapsed(opaque_phase);
    let graph = Instant::now();
    let input_metrics = package_metrics(&baseline)?;
    let output_metrics = package_metrics(&output)?;
    let graph_ns = elapsed(graph);
    Ok(success_result(
        fixture,
        baseline,
        output,
        input_members,
        output_members,
        output_metrics,
        input_metrics,
        semantic_ok,
        opaque_ok,
        true,
        PhaseTimes {
            capture_ns,
            publish_ns,
            opaque_ns,
            graph_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
    ))
}

fn run_signed_changed(fixture: &Fixture) -> Result<RunResult> {
    let mut package = open_package(fixture)?;
    if !package.is_signed() {
        return Err("signed fixture did not report a signature".into());
    }
    let baseline = package_bytes(&mut package)?;
    let input_members = member_hashes(&baseline)?;
    let snapshot = package.styles_with_effects(Owner::MainDocument)?;
    let resource = changed_resource(snapshot.resource().ok_or("signed main effects missing")?)?;
    let stage = Instant::now();
    let error = package
        .put_styles_with_effects(Owner::MainDocument, resource)
        .expect_err("changed signed publication unexpectedly succeeded");
    let stage_ns = elapsed(stage);
    let validation = Instant::now();
    let error_receipt = classify_error(&error, ExpectedError::Signed);
    let validation_ns = elapsed(validation);
    let (physical, metadata, readback_ns) = readback_unchanged(
        &mut package,
        &baseline,
        Owner::MainDocument,
        snapshot.resource(),
    )?;
    let graph = Instant::now();
    let metrics = package_metrics(&baseline)?;
    let graph_ns = elapsed(graph);
    Ok(RunResult {
        actual_success: false,
        ingress_refusal: false,
        no_output_ok: true,
        semantic_ok: error_receipt.typed_match && physical && metadata,
        opaque_ok: metadata,
        exact_inverse_ok: false,
        source_readback_physical_ok: Some(physical),
        source_readback_metadata_ok: Some(metadata),
        output_bytes: 0,
        input_sha256: sha256_hex(&baseline),
        output_sha256: None,
        input_metrics: metrics,
        output_metrics: None,
        input_member_digest: member_digest(&input_members),
        output_member_digest: None,
        cap_exact_fit_ok: None,
        cap_under_refused_ok: None,
        cap_refusal: None,
        cap_commit_refusal: None,
        cap_source_metrics: None,
        cap_projected_metrics: None,
        cap_existing: None,
        phases: PhaseTimes {
            stage_ns,
            validation_ns,
            graph_ns,
            readback_ns,
            ..PhaseTimes::default()
        },
        error: Some(error_receipt),
    })
}

fn run_independent(fixture: &Fixture, owner: Owner) -> Result<RunResult> {
    let capture = Instant::now();
    let mut package = open_package(fixture)?;
    let baseline = package_bytes(&mut package)?;
    let input_members = member_hashes(&baseline)?;
    let other = match owner {
        Owner::MainDocument => Owner::Glossary,
        Owner::Glossary => Owner::MainDocument,
    };
    let other_before = package
        .styles_with_effects(other)?
        .resource()
        .map(|resource| resource.xml_bytes().to_vec());
    let snapshot = package.styles_with_effects(owner)?;
    let capture_ns = elapsed(capture);
    let replacement = changed_resource(snapshot.resource().ok_or("independence owner absent")?)?;
    let stage = Instant::now();
    let mut edit = snapshot.edit();
    edit.replace_resource(Some(replacement.clone()))?;
    let stage_ns = elapsed(stage);
    let commit_start = Instant::now();
    let commit = edit.commit()?;
    let commit_ns = elapsed(commit_start);
    let publish = Instant::now();
    package.apply_styles_with_effects_patch(owner, commit.patch())?;
    let publish_ns = elapsed(publish);
    let output = package_bytes(&mut package)?;
    let reopen_start = Instant::now();
    let reopened = Package::from_reader(Cursor::new(output.as_slice()))?;
    let reopened_owner = reopened.styles_with_effects(owner)?;
    let other_after = reopened
        .styles_with_effects(other)?
        .resource()
        .map(|resource| resource.xml_bytes().to_vec());
    let reopen_ns = elapsed(reopen_start);
    let validation = Instant::now();
    let semantic_ok = reopened_owner
        .resource()
        .is_some_and(|resource| resource.xml_bytes() == replacement.xml_bytes())
        && other_before == other_after;
    let validation_ns = elapsed(validation);
    let output_members = member_hashes(&output)?;
    let opaque_phase = Instant::now();
    let opaque_ok = unchanged_except(&input_members, &output_members, &[effects_member(owner)]);
    let opaque_ns = elapsed(opaque_phase);
    let graph = Instant::now();
    let input_metrics = package_metrics(&baseline)?;
    let output_metrics = package_metrics(&output)?;
    let graph_ns = elapsed(graph);
    Ok(success_result(
        fixture,
        baseline,
        output,
        input_members,
        output_members,
        output_metrics,
        input_metrics,
        semantic_ok,
        opaque_ok,
        true,
        PhaseTimes {
            capture_ns,
            stage_ns,
            commit_ns,
            publish_ns,
            reopen_ns,
            opaque_ns,
            graph_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
    ))
}

#[derive(Clone, Copy)]
enum CapKind {
    Parts,
    TotalPartBytes,
    TotalRelationships,
    TotalRelationshipXmlEvents,
    TotalRelationshipXmlBytes,
    RelationshipParts,
    RelationshipGraphNodes,
}

fn cap_from_lane(lane: &str) -> Result<CapKind> {
    match lane {
        "cap_parts" => Ok(CapKind::Parts),
        "cap_total_part_bytes" => Ok(CapKind::TotalPartBytes),
        "cap_total_relationships" => Ok(CapKind::TotalRelationships),
        "cap_total_relationship_xml_events" => Ok(CapKind::TotalRelationshipXmlEvents),
        "cap_total_relationship_xml_bytes" => Ok(CapKind::TotalRelationshipXmlBytes),
        "cap_relationship_parts" => Ok(CapKind::RelationshipParts),
        "cap_relationship_graph_nodes" => Ok(CapKind::RelationshipGraphNodes),
        _ => Err(format!("unknown cap lane: {lane}").into()),
    }
}

fn run_cap(fixture: &Fixture, kind: CapKind) -> Result<RunResult> {
    let mut generous = open_package(fixture)?;
    let baseline = package_bytes(&mut generous)?;
    let input_members = member_hashes(&baseline)?;
    let resource = fixture_resource(COMPLEX, "word/stylesWithEffects.xml")?;
    let projected = projected_addition(&baseline, &resource)?;
    let source_value = cap_value(kind, &package_metrics(&baseline)?);
    let projected_value = cap_value(kind, &projected);
    if projected_value <= source_value {
        return Err(format!("cap did not grow: {source_value} -> {projected_value}").into());
    }
    let exact_limits = limits_with_cap(kind, projected_value)?;
    let mut exact = Package::from_reader_with_limits(Cursor::new(baseline.clone()), exact_limits)?;
    let exact_result = exact.put_styles_with_effects(Owner::MainDocument, resource.clone())?;
    let exact_output = package_bytes(&mut exact)?;
    let exact_members = member_hashes(&exact_output)?;
    let exact_opaque = unchanged_except(
        &input_members,
        &exact_members,
        &[
            effects_member(Owner::MainDocument),
            owner_relationship_member(Owner::MainDocument),
            target_relationship_member(Owner::MainDocument),
            "[Content_Types].xml",
        ],
    );
    let exact_ok = exact_result
        && metrics_equal(&package_metrics(&exact_output)?, &projected)
        && exact_members.contains_key(effects_member(Owner::MainDocument))
        && exact_opaque;
    let exact_reopened = Package::from_reader(Cursor::new(exact_output.as_slice()))?;
    let exact_owner = exact_reopened.styles_with_effects(Owner::MainDocument)?;
    let exact_ok = exact_ok
        && exact_owner
            .resource()
            .is_some_and(|value| value.xml_bytes() == resource.xml_bytes());

    let under_limit = projected_value.saturating_sub(1);
    let under_limits = limits_with_cap(kind, under_limit)?;
    let mut under = Package::from_reader_with_limits(Cursor::new(baseline.clone()), under_limits)?;
    let before_under = package_bytes(&mut under)?;
    let under_error = under
        .put_styles_with_effects(Owner::MainDocument, resource.clone())
        .expect_err("one-unit-under cap unexpectedly succeeded");
    let expected_resource = cap_read_resource(kind);
    let validation = Instant::now();
    let under_receipt = classify_error(&under_error, ExpectedError::ReadLimit(expected_resource));
    let under_validation_ns = elapsed(validation);
    let under_readback = Instant::now();
    let after_under = package_bytes(&mut under)?;
    let physical = after_under == before_under;
    let reopened = Package::from_reader(Cursor::new(after_under.as_slice()))?;
    let metadata = reopened
        .styles_with_effects(Owner::MainDocument)?
        .is_empty();
    let under_metrics = package_metrics(&after_under)?;
    let under_members = member_hashes(&after_under)?;
    let baseline_members = member_hashes(&before_under)?;
    let opaque_unchanged = metrics_equal(&under_metrics, &package_metrics(&baseline)?)
        && under_members == baseline_members;
    let readback_ns = elapsed(under_readback);

    // The same cap must fail at Transaction::commit, before publication
    // metadata is serialized. This is a source-bound staging gate, not a
    // canned publication refusal.
    let mut staged_package =
        Package::from_reader_with_limits(Cursor::new(baseline.clone()), under_limits)?;
    let staged_snapshot = staged_package.styles_with_effects(Owner::MainDocument)?;
    let stage = Instant::now();
    let mut edit = staged_snapshot.edit();
    edit.replace_resource(Some(resource))?;
    let commit_error = edit
        .commit()
        .expect_err("one-unit-under cap unexpectedly passed transaction commit");
    let commit_ns = elapsed(stage);
    let validation = Instant::now();
    let commit_receipt = classify_error(&commit_error, ExpectedError::ReadLimit(expected_resource));
    let commit_validation_ns = elapsed(validation);
    let stage_source_unchanged = package_bytes(&mut staged_package)? == baseline;
    let semantic_ok = exact_ok
        && under_receipt.typed_match
        && commit_receipt.typed_match
        && physical
        && metadata
        && opaque_unchanged
        && stage_source_unchanged;
    let graph = Instant::now();
    let output_metrics = package_metrics(&exact_output)?;
    let input_metrics = package_metrics(&baseline)?;
    let output_member_digest = member_digest(&exact_members);
    let graph_ns = elapsed(graph);
    Ok(RunResult {
        actual_success: true,
        ingress_refusal: false,
        no_output_ok: true,
        semantic_ok,
        opaque_ok: exact_opaque && physical && metadata && opaque_unchanged,
        exact_inverse_ok: true,
        source_readback_physical_ok: Some(physical && stage_source_unchanged),
        source_readback_metadata_ok: Some(metadata),
        output_bytes: exact_output.len() as u64,
        input_sha256: sha256_hex(&baseline),
        output_sha256: Some(sha256_hex(&exact_output)),
        input_metrics,
        output_metrics: Some(output_metrics),
        input_member_digest: member_digest(&input_members),
        output_member_digest: Some(output_member_digest),
        cap_exact_fit_ok: Some(exact_ok),
        cap_under_refused_ok: Some(
            under_receipt.typed_match
                && commit_receipt.typed_match
                && physical
                && metadata
                && opaque_unchanged,
        ),
        cap_refusal: Some(under_receipt),
        cap_commit_refusal: Some(commit_receipt),
        cap_source_metrics: Some(input_metrics),
        cap_projected_metrics: Some(output_metrics),
        cap_existing: Some(existing_owner_cap_evidence(&native_bug_fixture(), kind)?),
        phases: PhaseTimes {
            commit_ns,
            readback_ns,
            graph_ns,
            validation_ns: under_validation_ns + commit_validation_ns,
            ..PhaseTimes::default()
        },
        error: None,
    })
}

fn malformed_opaque_fixture(source: &[u8]) -> Result<Vec<u8>> {
    rewrite_zip(
        source,
        |name, bytes| {
            if name == effects_member(Owner::MainDocument) {
                return Ok(Some(replace_once(
                    bytes,
                    b"</w:styles>",
                    b"&unknown;</w:styles>",
                )?));
            }
            Ok(Some(bytes.to_vec()))
        },
        &[],
    )
}

fn run_malformed_opaque_xml(fixture: &Fixture) -> Result<RunResult> {
    run_malformed_bytes(
        malformed_opaque_fixture(fixture.package.as_ref())?,
        ExpectedError::MalformedOpaque,
    )
}

fn malformed_xml_case_fixture(source: &[u8], lane: &str) -> Result<Vec<u8>> {
    let (attributes, body, prolog) = match lane {
        "malformed_unbound_descendant" => ("", "<x:future/>", ""),
        "malformed_invalid_qname" => ("", "<1bad/>", ""),
        "malformed_raw_attribute" => (" bad=\"<\"", "", ""),
        "malformed_raw_text" => ("", "]]>", ""),
        "malformed_control" => ("", "\u{1}", ""),
        "malformed_invalid_char_ref" => ("", "&#x1;", ""),
        "malformed_empty_prefix" => (" xmlns:x=\"\" x:y=\"z\"", "", ""),
        "malformed_reserved_xml_uri" => (" xmlns=\"http://www.w3.org/XML/1998/namespace\"", "", ""),
        "malformed_xml_version" => ("", "", "<?xml version=\"2.0\"?>"),
        _ => return Err(format!("unknown malformed XML lane: {lane}").into()),
    };
    let xml =
        format!("{prolog}<w:styles xmlns:w=\"{TRANSITIONAL_W}\"{attributes}>{body}</w:styles>");
    rewrite_zip(
        source,
        |name, bytes| {
            if name == effects_member(Owner::MainDocument) {
                return Ok(Some(xml.as_bytes().to_vec()));
            }
            Ok(Some(bytes.to_vec()))
        },
        &[],
    )
}

fn run_malformed_xml_case(fixture: &Fixture, lane: &str) -> Result<RunResult> {
    run_malformed_bytes(
        malformed_xml_case_fixture(fixture.package.as_ref(), lane)?,
        ExpectedError::Malformed,
    )
}

fn xml_shape(xml: &[u8]) -> Result<(usize, usize)> {
    let mut reader = Reader::from_reader(xml);
    let mut events = 0_usize;
    let mut depth = 0_usize;
    let mut maximum_depth = 0_usize;
    loop {
        events = events.checked_add(1).ok_or("XML event metric overflow")?;
        match reader.read_event()? {
            Event::Start(_) => {
                depth = depth.checked_add(1).ok_or("XML depth metric overflow")?;
                maximum_depth = maximum_depth.max(depth);
            },
            Event::End(_) => depth = depth.checked_sub(1).ok_or("XML depth underflow")?,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok((events, maximum_depth))
}

fn run_xml_limit_refusal(fixture: &Fixture, depth_limit: bool) -> Result<RunResult> {
    let source = fixture.package.as_ref();
    let xml = member_bytes(source)?
        .remove(effects_member(Owner::MainDocument))
        .ok_or("XML-limit fixture effects member missing")?;
    let (events, depth) = xml_shape(&xml)?;
    let (limit, expected_resource) = if depth_limit {
        (
            ReadLimits::builder()
                .max_xml_depth(depth.saturating_sub(1))?
                .build()?,
            ReadResource::XmlDepth,
        )
    } else {
        (
            ReadLimits::builder()
                .max_xml_events(events.saturating_sub(1))?
                .build()?,
            ReadResource::XmlEvents,
        )
    };
    let baseline = source.to_vec();
    let input_metrics = package_metrics(&baseline)?;
    let input_member_digest = member_digest(&member_hashes(&baseline)?);
    let capture = Instant::now();
    let mut package = match Package::from_reader_with_limits(Cursor::new(baseline.clone()), limit) {
        Ok(package) => package,
        Err(error) => {
            let capture_ns = elapsed(capture);
            let validation = Instant::now();
            let receipt = classify_error(&error, ExpectedError::ReadLimit(expected_resource));
            let validation_ns = elapsed(validation);
            return Ok(RunResult {
                actual_success: false,
                ingress_refusal: true,
                no_output_ok: true,
                semantic_ok: receipt.typed_match,
                opaque_ok: receipt.typed_match,
                exact_inverse_ok: false,
                source_readback_physical_ok: None,
                source_readback_metadata_ok: None,
                output_bytes: 0,
                input_sha256: sha256_hex(&baseline),
                output_sha256: None,
                input_metrics,
                output_metrics: None,
                input_member_digest,
                output_member_digest: None,
                cap_exact_fit_ok: None,
                cap_under_refused_ok: None,
                cap_refusal: None,
                cap_commit_refusal: None,
                cap_source_metrics: None,
                cap_projected_metrics: None,
                cap_existing: None,
                phases: PhaseTimes {
                    capture_ns,
                    validation_ns,
                    ..PhaseTimes::default()
                },
                error: Some(receipt),
            });
        },
    };
    let error = package
        .styles_with_effects(Owner::MainDocument)
        .expect_err("caller XML limit unexpectedly accepted native effects XML");
    let capture_ns = elapsed(capture);
    let validation = Instant::now();
    let receipt = classify_error(&error, ExpectedError::ReadLimit(expected_resource));
    let validation_ns = elapsed(validation);
    let readback = Instant::now();
    let after = package_bytes(&mut package)?;
    let physical = after == baseline;
    let reopened = Package::from_reader(Cursor::new(after.as_slice()))?;
    let metadata = reopened
        .styles_with_effects(Owner::MainDocument)?
        .resource()
        .is_some();
    let readback_ns = elapsed(readback);
    Ok(RunResult {
        actual_success: false,
        ingress_refusal: false,
        no_output_ok: true,
        semantic_ok: receipt.typed_match && physical && metadata,
        opaque_ok: receipt.typed_match && physical && metadata,
        exact_inverse_ok: false,
        source_readback_physical_ok: Some(physical),
        source_readback_metadata_ok: Some(metadata),
        output_bytes: 0,
        input_sha256: sha256_hex(&baseline),
        output_sha256: None,
        input_metrics,
        output_metrics: None,
        input_member_digest,
        output_member_digest: None,
        cap_exact_fit_ok: None,
        cap_under_refused_ok: None,
        cap_refusal: None,
        cap_commit_refusal: None,
        cap_source_metrics: None,
        cap_projected_metrics: None,
        cap_existing: None,
        phases: PhaseTimes {
            capture_ns,
            readback_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
        error: Some(receipt),
    })
}

fn run_malformed(fixture: &Fixture, lane: &str) -> Result<RunResult> {
    run_malformed_bytes(
        malformed_fixture(fixture.package.as_ref(), lane)?,
        ExpectedError::Malformed,
    )
}

fn run_malformed_bytes(malformed: Vec<u8>, expected: ExpectedError) -> Result<RunResult> {
    let input_sha256 = sha256_hex(&malformed);
    let capture = Instant::now();
    let mut package = match Package::from_reader(Cursor::new(malformed.as_slice())) {
        Ok(package) => package,
        Err(error) => {
            let capture_ns = elapsed(capture);
            let validation = Instant::now();
            let receipt = classify_error(&error, expected);
            let validation_ns = elapsed(validation);
            let graph = Instant::now();
            let input_metrics = package_metrics(&malformed)?;
            let input_member_digest = member_digest(&member_hashes(&malformed)?);
            let graph_ns = elapsed(graph);
            return Ok(RunResult {
                actual_success: false,
                ingress_refusal: true,
                no_output_ok: true,
                semantic_ok: receipt.typed_match,
                opaque_ok: receipt.typed_match,
                exact_inverse_ok: false,
                source_readback_physical_ok: None,
                source_readback_metadata_ok: None,
                output_bytes: 0,
                input_sha256,
                output_sha256: None,
                input_metrics,
                output_metrics: None,
                input_member_digest,
                output_member_digest: None,
                cap_exact_fit_ok: None,
                cap_under_refused_ok: None,
                cap_refusal: None,
                cap_commit_refusal: None,
                cap_source_metrics: None,
                cap_projected_metrics: None,
                cap_existing: None,
                phases: PhaseTimes {
                    capture_ns,
                    graph_ns,
                    validation_ns,
                    ..PhaseTimes::default()
                },
                error: Some(receipt),
            });
        },
    };
    let before = package_bytes(&mut package)?;
    let error = match package.styles_with_effects(Owner::MainDocument) {
        Ok(_) => {
            let capture_ns = elapsed(capture);
            let graph = Instant::now();
            let input_metrics = package_metrics(&malformed)?;
            let input_member_digest = member_digest(&member_hashes(&malformed)?);
            let graph_ns = elapsed(graph);
            return Ok(RunResult {
                actual_success: false,
                ingress_refusal: false,
                no_output_ok: true,
                semantic_ok: false,
                opaque_ok: false,
                exact_inverse_ok: false,
                source_readback_physical_ok: None,
                source_readback_metadata_ok: None,
                output_bytes: 0,
                input_sha256,
                output_sha256: None,
                input_metrics,
                output_metrics: None,
                input_member_digest,
                output_member_digest: None,
                cap_exact_fit_ok: None,
                cap_under_refused_ok: None,
                cap_refusal: None,
                cap_commit_refusal: None,
                cap_source_metrics: None,
                cap_projected_metrics: None,
                cap_existing: None,
                phases: PhaseTimes {
                    capture_ns,
                    graph_ns,
                    ..PhaseTimes::default()
                },
                error: Some(ErrorReceipt {
                    class: "unexpected_success".to_owned(),
                    variant: "unexpected_success".to_owned(),
                    message: "malformed topology was accepted by the public API".to_owned(),
                    typed_match: false,
                    resource: None,
                    actual: None,
                    maximum: None,
                }),
            });
        },
        Err(error) => error,
    };
    let capture_ns = elapsed(capture);
    let validation = Instant::now();
    let receipt = classify_error(&error, expected);
    let validation_ns = elapsed(validation);
    let readback = Instant::now();
    let after = package_bytes(&mut package)?;
    let physical = after == before;
    let member_equal = member_hashes(&after)? == member_hashes(&before)?;
    let metadata = member_equal
        && Package::from_reader(Cursor::new(after.as_slice()))
            .and_then(|reopened| {
                reopened
                    .styles_with_effects(Owner::MainDocument)
                    .map(|_| ())
            })
            .is_err();
    let readback_ns = elapsed(readback);
    let graph = Instant::now();
    let input_metrics = package_metrics(&malformed)?;
    let input_member_digest = member_digest(&member_hashes(&malformed)?);
    let graph_ns = elapsed(graph);
    Ok(RunResult {
        actual_success: false,
        ingress_refusal: false,
        no_output_ok: true,
        semantic_ok: receipt.typed_match && physical && metadata,
        opaque_ok: receipt.typed_match && physical && metadata,
        exact_inverse_ok: false,
        source_readback_physical_ok: Some(physical),
        source_readback_metadata_ok: Some(metadata),
        output_bytes: 0,
        input_sha256,
        output_sha256: None,
        input_metrics,
        output_metrics: None,
        input_member_digest,
        output_member_digest: None,
        cap_exact_fit_ok: None,
        cap_under_refused_ok: None,
        cap_refusal: None,
        cap_commit_refusal: None,
        cap_source_metrics: None,
        cap_projected_metrics: None,
        cap_existing: None,
        phases: PhaseTimes {
            capture_ns,
            graph_ns,
            readback_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
        error: Some(receipt),
    })
}

#[derive(Clone, Copy)]
enum ExpectedError {
    ReadLimit(ReadResource),
    Signed,
    GlossaryMissing,
    StalePatch,
    Malformed,
    MalformedOpaque,
}

fn classify_error(error: &DocxError, expected: ExpectedError) -> ErrorReceipt {
    let (variant, resource, actual, maximum) = match error {
        DocxError::Opc(OpcError::ReadLimit {
            resource,
            actual,
            maximum,
        }) => (
            "DocxError::Opc::ReadLimit",
            Some(*resource),
            Some(*actual),
            Some(*maximum),
        ),
        DocxError::Opc(OpcError::SignedSourceRequiresExplicitPolicy) => (
            "DocxError::Opc::SignedSourceRequiresExplicitPolicy",
            None,
            None,
            None,
        ),
        DocxError::Opc(OpcError::SourceBackedOverlayUnavailable { .. }) => (
            "DocxError::Opc::SourceBackedOverlayUnavailable",
            None,
            None,
            None,
        ),
        DocxError::PartNotFound(_) => ("DocxError::PartNotFound", None, None, None),
        DocxError::InvalidFormat(_) => ("DocxError::InvalidFormat", None, None, None),
        DocxError::InvalidRelationship(_) => ("DocxError::InvalidRelationship", None, None, None),
        DocxError::ContentType { .. } => ("DocxError::ContentType", None, None, None),
        DocxError::InvalidContentType { .. } => ("DocxError::InvalidContentType", None, None, None),
        DocxError::Xml(_) => ("DocxError::Xml", None, None, None),
        DocxError::Invalid(_) => ("DocxError::Invalid", None, None, None),
        DocxError::UnsafeEdit { .. } => ("DocxError::UnsafeEdit", None, None, None),
        DocxError::Opc(_) => ("DocxError::Opc", None, None, None),
        _ => ("DocxError::Other", None, None, None),
    };
    let typed_match = match expected {
        ExpectedError::ReadLimit(wanted) => {
            resource == Some(wanted)
                && actual
                    .zip(maximum)
                    .is_some_and(|(actual, maximum)| actual > maximum)
        },
        ExpectedError::Signed => variant == "DocxError::UnsafeEdit",
        ExpectedError::GlossaryMissing => variant == "DocxError::PartNotFound",
        ExpectedError::StalePatch => variant == "DocxError::InvalidFormat",
        ExpectedError::Malformed => variant != "DocxError::Other",
        ExpectedError::MalformedOpaque => {
            variant == "DocxError::Opc::SourceBackedOverlayUnavailable"
        },
    };
    ErrorReceipt {
        class: variant.to_owned(),
        variant: variant.to_owned(),
        message: error.to_string(),
        typed_match,
        resource: resource.map(|value| format!("{value:?}")),
        actual,
        maximum,
    }
}

fn readback_unchanged(
    package: &mut Package,
    baseline: &[u8],
    owner: Owner,
    expected_resource: Option<&Resource>,
) -> Result<(bool, bool, u64)> {
    let readback = Instant::now();
    let after = package_bytes(package)?;
    let physical = after == baseline;
    let member_equal = member_hashes(&after)? == member_hashes(baseline)?;
    let reopened = Package::from_reader(Cursor::new(after.as_slice()))?;
    let snapshot = reopened.styles_with_effects(owner)?;
    let metadata = member_equal
        && match (expected_resource, snapshot.resource()) {
            (None, None) => true,
            (Some(expected), Some(actual)) => {
                expected.xml_bytes() == actual.xml_bytes()
                    && expected.conformance() == actual.conformance()
            },
            _ => false,
        };
    Ok((physical, metadata, elapsed(readback)))
}

fn success_result(
    _fixture: &Fixture,
    input: Vec<u8>,
    output: Vec<u8>,
    input_members: BTreeMap<String, MemberDigest>,
    output_members: BTreeMap<String, MemberDigest>,
    output_metrics: PackageMetrics,
    input_metrics: PackageMetrics,
    semantic_ok: bool,
    opaque_ok: bool,
    exact_inverse_ok: bool,
    phases: PhaseTimes,
) -> RunResult {
    RunResult {
        actual_success: true,
        ingress_refusal: false,
        no_output_ok: true,
        semantic_ok,
        opaque_ok,
        exact_inverse_ok,
        source_readback_physical_ok: None,
        source_readback_metadata_ok: None,
        output_bytes: output.len() as u64,
        input_sha256: sha256_hex(&input),
        output_sha256: Some(sha256_hex(&output)),
        input_metrics,
        output_metrics: Some(output_metrics),
        input_member_digest: member_digest(&input_members),
        output_member_digest: Some(member_digest(&output_members)),
        cap_exact_fit_ok: None,
        cap_under_refused_ok: None,
        cap_refusal: None,
        cap_commit_refusal: None,
        cap_source_metrics: None,
        cap_projected_metrics: None,
        cap_existing: None,
        phases,
        error: None,
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MemberDigest {
    pub length: u64,
    pub sha256: String,
}

fn member_hashes(bytes: &[u8]) -> Result<BTreeMap<String, MemberDigest>> {
    let members = member_bytes(bytes)?;
    Ok(members
        .into_iter()
        .map(|(name, bytes)| {
            (
                name,
                MemberDigest {
                    length: bytes.len() as u64,
                    sha256: sha256_hex(&bytes),
                },
            )
        })
        .collect())
}

fn member_bytes(bytes: &[u8]) -> Result<BTreeMap<String, Vec<u8>>> {
    let reader = PhysPkgReader::new(bytes)?;
    let mut output = BTreeMap::new();
    for name in reader.member_names()? {
        output.insert(name.clone(), reader.read_member(&name)?);
    }
    Ok(output)
}

fn member_digest(members: &BTreeMap<String, MemberDigest>) -> String {
    let mut hasher = Sha256::new();
    for (name, digest) in members {
        hasher.update(name.as_bytes());
        hasher.update([0]);
        hasher.update(digest.length.to_le_bytes());
        hasher.update(digest.sha256.as_bytes());
        hasher.update([0]);
    }
    hex_bytes(&hasher.finalize())
}

fn unchanged_except(
    before: &BTreeMap<String, MemberDigest>,
    after: &BTreeMap<String, MemberDigest>,
    allowed: &[&str],
) -> bool {
    let allowed = allowed.iter().copied().collect::<BTreeSet<_>>();
    let names = before
        .keys()
        .chain(after.keys())
        .cloned()
        .collect::<BTreeSet<_>>();
    names
        .into_iter()
        .all(|name| allowed.contains(name.as_str()) || before.get(&name) == after.get(&name))
}

fn unrelated_member_digest(members: &BTreeMap<String, MemberDigest>, allowed: &[&str]) -> String {
    let allowed = allowed.iter().copied().collect::<BTreeSet<_>>();
    let mut hasher = Sha256::new();
    for (name, digest) in members {
        if allowed.contains(name.as_str()) {
            continue;
        }
        hasher.update(name.as_bytes());
        hasher.update([0]);
        hasher.update(digest.length.to_le_bytes());
        hasher.update(digest.sha256.as_bytes());
        hasher.update([0]);
    }
    hex_bytes(&hasher.finalize())
}

fn package_metrics(bytes: &[u8]) -> Result<PackageMetrics> {
    let opc = OpcPackage::from_vec(bytes.to_vec())?;
    let mut total_part_bytes = 0_u64;
    let mut total_relationships = opc.rels().iter().count() as u64;
    for part in opc.iter_parts() {
        total_part_bytes = total_part_bytes
            .checked_add(part.blob().len() as u64)
            .ok_or("part byte metric overflow")?;
        total_relationships = total_relationships
            .checked_add(part.rels().iter().count() as u64)
            .ok_or("relationship metric overflow")?;
    }
    let reader = PhysPkgReader::new(bytes)?;
    let mut relationship_parts = 0_u64;
    let mut relationship_xml_bytes = 0_u64;
    let mut relationship_xml_events = 0_u64;
    for name in reader.member_names()? {
        if is_relationship_member(&name) {
            relationship_parts += 1;
            let xml = reader.read_member(&name)?;
            relationship_xml_bytes += xml.len() as u64;
            relationship_xml_events += xml_event_count(&xml)?;
        }
    }
    let relationship_graph_nodes = relationship_graph_nodes(&opc)? as u64;
    Ok(PackageMetrics {
        parts: opc.part_count() as u64,
        total_part_bytes,
        total_relationships,
        relationship_parts,
        relationship_graph_nodes,
        relationship_xml_bytes,
        relationship_xml_events,
    })
}

fn relationship_graph_nodes(package: &OpcPackage) -> Result<usize> {
    let mut visited = BTreeSet::new();
    let mut queue = VecDeque::new();
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        queue.push_back(relationship.target_partname()?);
    }
    while let Some(source) = queue.pop_front() {
        let key = source.as_str().to_owned();
        if !visited.insert(key) {
            continue;
        }
        let Ok(part) = package.get_part(&source) else {
            continue;
        };
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            queue.push_back(relationship.target_partname()?);
        }
    }
    Ok(visited.len())
}

fn xml_event_count(bytes: &[u8]) -> Result<u64> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut count = 0_u64;
    loop {
        count = count.checked_add(1).ok_or("XML event metric overflow")?;
        if matches!(reader.read_event_into(&mut buffer)?, Event::Eof) {
            break;
        }
        buffer.clear();
    }
    Ok(count)
}

fn check_native_expectation(fixture: &Fixture) -> Result<()> {
    let expectation = native_expectation(fixture.name)
        .ok_or_else(|| format!("no native expectation for {}", fixture.name))?;
    let actual = sha256_hex(fixture.package.as_ref());
    if actual != expectation.package_sha256 {
        return Err(format!("native package hash changed for {}", fixture.name).into());
    }
    let members = member_bytes(fixture.package.as_ref())?;
    for expected in [expectation.main, expectation.glossary]
        .into_iter()
        .flatten()
    {
        let bytes = members
            .get(expected.name)
            .ok_or_else(|| format!("native member missing: {}", expected.name))?;
        if bytes.len() != expected.length || sha256_hex(bytes) != expected.sha256 {
            return Err(format!("native member hash/length changed: {}", expected.name).into());
        }
    }
    Ok(())
}

fn fixture_resource(bytes: &[u8], member: &str) -> Result<Resource> {
    let xml = member_bytes(bytes)?
        .remove(member)
        .ok_or_else(|| format!("resource member missing: {member}"))?;
    Ok(Resource::from_xml(xml)?)
}

fn changed_resource(resource: &Resource) -> Result<Resource> {
    let marker = br#"<w:styles"#;
    let closing = br#"</w:styles>"#;
    let xml = resource.xml_bytes();
    let offset = xml
        .windows(closing.len())
        .position(|window| window == closing)
        .ok_or("effects XML has no styles closing element")?;
    let mut changed = Vec::with_capacity(xml.len() + 72);
    changed.extend_from_slice(&xml[..offset]);
    changed.extend_from_slice(
        br#"<w:extensibilityMarker xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" data="docx-effects-smoke-v1"/>"#,
    );
    changed.extend_from_slice(&xml[offset..]);
    if !xml.windows(marker.len()).any(|window| window == marker) {
        return Err("effects XML does not have the expected w:styles root".into());
    }
    let replacement = Resource::from_xml(changed)?;
    if replacement.conformance() != resource.conformance() {
        return Err("replacement conformance changed".into());
    }
    Ok(replacement)
}

fn snapshot_matches_package(
    bytes: &[u8],
    owner: Owner,
    expected: Option<&Resource>,
) -> Result<bool> {
    let package = Package::from_reader(Cursor::new(bytes))?;
    let actual = package.styles_with_effects(owner)?;
    Ok(match (expected, actual.resource()) {
        (None, None) => true,
        (Some(expected), Some(actual)) => {
            expected.xml_bytes() == actual.xml_bytes()
                && expected.conformance() == actual.conformance()
        },
        _ => false,
    })
}

fn open_package(fixture: &Fixture) -> Result<Package> {
    Ok(Package::from_reader(Cursor::new(fixture.package.as_ref()))?)
}

fn package_bytes(package: &mut Package) -> Result<Vec<u8>> {
    let mut output = Cursor::new(Vec::new());
    package.to_stream(&mut output)?;
    Ok(output.into_inner())
}

fn elapsed(start: Instant) -> u64 {
    start.elapsed().as_nanos().try_into().unwrap_or(u64::MAX)
}

fn sha256_hex(bytes: &[u8]) -> String {
    hex_bytes(&Sha256::digest(bytes))
}

fn hex_bytes(bytes: &[u8]) -> String {
    let mut output = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn effects_member(owner: Owner) -> &'static str {
    match owner {
        Owner::MainDocument => "word/stylesWithEffects.xml",
        Owner::Glossary => "word/glossary/stylesWithEffects.xml",
    }
}

fn owner_relationship_member(owner: Owner) -> &'static str {
    match owner {
        Owner::MainDocument => "word/_rels/document.xml.rels",
        Owner::Glossary => "word/glossary/_rels/document.xml.rels",
    }
}

fn target_relationship_member(owner: Owner) -> &'static str {
    match owner {
        Owner::MainDocument => "word/_rels/stylesWithEffects.xml.rels",
        Owner::Glossary => "word/glossary/_rels/stylesWithEffects.xml.rels",
    }
}

fn owner_edit_members(owner: Owner) -> [&'static str; 4] {
    [
        effects_member(owner),
        owner_relationship_member(owner),
        target_relationship_member(owner),
        "[Content_Types].xml",
    ]
}

fn is_relationship_member(name: &str) -> bool {
    name == "_rels/.rels" || (name.contains("/_rels/") && name.ends_with(".rels"))
}

fn cap_value(kind: CapKind, metrics: &PackageMetrics) -> u64 {
    match kind {
        CapKind::Parts => metrics.parts,
        CapKind::TotalPartBytes => metrics.total_part_bytes,
        CapKind::TotalRelationships => metrics.total_relationships,
        CapKind::TotalRelationshipXmlEvents => metrics.relationship_xml_events,
        CapKind::TotalRelationshipXmlBytes => metrics.relationship_xml_bytes,
        CapKind::RelationshipParts => metrics.relationship_parts,
        CapKind::RelationshipGraphNodes => metrics.relationship_graph_nodes,
    }
}

fn cap_read_resource(kind: CapKind) -> ReadResource {
    match kind {
        CapKind::Parts => ReadResource::Parts,
        CapKind::TotalPartBytes => ReadResource::TotalPartBytes,
        CapKind::TotalRelationships => ReadResource::TotalRelationships,
        CapKind::TotalRelationshipXmlEvents => ReadResource::TotalRelationshipXmlEvents,
        CapKind::TotalRelationshipXmlBytes => ReadResource::TotalRelationshipXmlBytes,
        CapKind::RelationshipParts => ReadResource::RelationshipParts,
        CapKind::RelationshipGraphNodes => ReadResource::RelationshipGraphNodes,
    }
}

fn limits_with_cap(kind: CapKind, value: u64) -> Result<ReadLimits> {
    let builder = ReadLimits::builder();
    let builder = match kind {
        CapKind::Parts => builder.max_parts(usize::try_from(value)?)?,
        CapKind::TotalPartBytes => builder.max_total_part_bytes(value)?,
        CapKind::TotalRelationships => builder.max_total_relationships(usize::try_from(value)?)?,
        CapKind::TotalRelationshipXmlEvents => {
            builder.max_total_relationship_xml_events(usize::try_from(value)?)?
        },
        CapKind::TotalRelationshipXmlBytes => {
            builder.max_total_relationship_xml_bytes(usize::try_from(value)?)?
        },
        CapKind::RelationshipParts => builder.max_relationship_parts(usize::try_from(value)?)?,
        CapKind::RelationshipGraphNodes => {
            builder.max_relationship_graph_nodes(usize::try_from(value)?)?
        },
    };
    Ok(builder.build()?)
}

fn metrics_equal(left: &PackageMetrics, right: &PackageMetrics) -> bool {
    left.parts == right.parts
        && left.total_part_bytes == right.total_part_bytes
        && left.total_relationships == right.total_relationships
        && left.relationship_parts == right.relationship_parts
        && left.relationship_graph_nodes == right.relationship_graph_nodes
        && left.relationship_xml_bytes == right.relationship_xml_bytes
        && left.relationship_xml_events == right.relationship_xml_events
}

fn native_bug_fixture() -> Fixture {
    Fixture {
        name: "Bug54849.docx",
        package: Arc::from(BUG),
        native: true,
        signed: false,
        main_present: true,
        glossary_present: true,
        expected_package_sha256: Some(
            "f54182713ea5ce5d77b9593d3d9d24e645460043cec0b40ef59c932385f084d3",
        ),
    }
}

fn existing_owner_cap_evidence(fixture: &Fixture, kind: CapKind) -> Result<CapEvidence> {
    let mut generous = open_package(fixture)?;
    let baseline = package_bytes(&mut generous)?;
    let source_members = member_hashes(&baseline)?;
    let allowed_members = owner_edit_members(Owner::MainDocument);
    let source_metrics = package_metrics(&baseline)?;
    let snapshot = generous.styles_with_effects(Owner::MainDocument)?;
    let replacement = changed_resource(
        snapshot
            .resource()
            .ok_or("existing-owner cap source is absent")?,
    )?;
    generous.put_styles_with_effects(Owner::MainDocument, replacement.clone())?;
    let projected_output = package_bytes(&mut generous)?;
    let projected_metrics = package_metrics(&projected_output)?;
    let source_value = cap_value(kind, &source_metrics);
    let projected_value = cap_value(kind, &projected_metrics);
    if projected_value <= source_value {
        return Ok(CapEvidence {
            applicable: false,
            source_metrics,
            projected_metrics,
            exact_fit_ok: false,
            exact_opaque_ok: false,
            source_unrelated_member_digest: None,
            exact_unrelated_member_digest: None,
            under_refused_ok: false,
            commit_stage_checked: false,
            refusal: None,
            commit_refusal: None,
        });
    }

    let exact_limits = limits_with_cap(kind, projected_value)?;
    let mut exact = Package::from_reader_with_limits(Cursor::new(baseline.clone()), exact_limits)?;
    let exact_before = package_bytes(&mut exact)?;
    let exact_changed = exact.put_styles_with_effects(Owner::MainDocument, replacement.clone())?;
    let exact_output = package_bytes(&mut exact)?;
    let exact_members = member_hashes(&exact_output)?;
    let exact_metrics = package_metrics(&exact_output)?;
    let exact_reopened = Package::from_reader(Cursor::new(exact_output.as_slice()))?;
    let exact_owner = exact_reopened.styles_with_effects(Owner::MainDocument)?;
    let exact_opaque_ok = unchanged_except(&source_members, &exact_members, &allowed_members);
    let exact_fit_ok = exact_before == baseline
        && exact_changed
        && metrics_equal(&exact_metrics, &projected_metrics)
        && exact_opaque_ok
        && exact_owner
            .resource()
            .is_some_and(|resource| resource.xml_bytes() == replacement.xml_bytes());
    let source_unrelated_member_digest = unrelated_member_digest(&source_members, &allowed_members);
    let exact_unrelated_member_digest = unrelated_member_digest(&exact_members, &allowed_members);

    let under_limits = limits_with_cap(kind, projected_value - 1)?;
    let mut under = Package::from_reader_with_limits(Cursor::new(baseline.clone()), under_limits)?;
    let before_under = package_bytes(&mut under)?;
    let under_error = under
        .put_styles_with_effects(Owner::MainDocument, replacement.clone())
        .expect_err("existing-owner one-unit-under cap unexpectedly succeeded");
    let under_receipt = classify_error(
        &under_error,
        ExpectedError::ReadLimit(cap_read_resource(kind)),
    );
    let after_under = package_bytes(&mut under)?;
    let under_physical = after_under == before_under && after_under == baseline;
    let under_reopened = Package::from_reader(Cursor::new(after_under.as_slice()))?;
    let under_metadata = under_reopened
        .styles_with_effects(Owner::MainDocument)?
        .resource()
        .is_some_and(|resource| {
            snapshot.resource().is_some_and(|base| {
                resource.xml_bytes() == base.xml_bytes()
                    && resource.conformance() == base.conformance()
            })
        });

    let mut staged = Package::from_reader_with_limits(Cursor::new(baseline), under_limits)?;
    let staged_snapshot = staged.styles_with_effects(Owner::MainDocument)?;
    let mut edit = staged_snapshot.edit();
    edit.replace_resource(Some(replacement))?;
    let (commit_stage_checked, commit_receipt) = match edit.commit() {
        Ok(_) => (false, None),
        Err(error) => (
            true,
            Some(classify_error(
                &error,
                ExpectedError::ReadLimit(cap_read_resource(kind)),
            )),
        ),
    };
    let staged_unchanged = package_bytes(&mut staged)? == before_under;
    Ok(CapEvidence {
        applicable: true,
        source_metrics,
        projected_metrics,
        exact_fit_ok,
        exact_opaque_ok,
        source_unrelated_member_digest: Some(source_unrelated_member_digest),
        exact_unrelated_member_digest: Some(exact_unrelated_member_digest),
        under_refused_ok: under_receipt.typed_match
            && under_physical
            && under_metadata
            && staged_unchanged,
        commit_stage_checked,
        refusal: Some(under_receipt),
        commit_refusal: commit_receipt,
    })
}

fn projected_addition(source: &[u8], resource: &Resource) -> Result<PackageMetrics> {
    let mut package = Package::from_reader(Cursor::new(source))?;
    package.put_styles_with_effects(Owner::MainDocument, resource.clone())?;
    package_metrics(&package_bytes(&mut package)?)
}

fn remove_main_effects(source: &[u8]) -> Result<Vec<u8>> {
    rewrite_zip(
        source,
        |name, bytes| {
            if name == "word/stylesWithEffects.xml" {
                return Ok(None);
            }
            let mut bytes = bytes.to_vec();
            if name == "word/_rels/document.xml.rels" {
                bytes = replace_once(
                &bytes,
                br#"<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects" Target="stylesWithEffects.xml"/>"#,
                b"",
            )?;
            }
            if name == "[Content_Types].xml" {
                bytes = replace_once(
                &bytes,
                br#"<Override PartName="/word/stylesWithEffects.xml" ContentType="application/vnd.ms-word.stylesWithEffects+xml"/>"#,
                b"",
            )?;
            }
            Ok(Some(bytes))
        },
        &[],
    )
}

fn malformed_fixture(source: &[u8], lane: &str) -> Result<Vec<u8>> {
    let mut additions = Vec::new();
    let output = match lane {
        "malformed_duplicate_owner" => rewrite_zip(
            source,
            |name, bytes| {
                if name == "word/_rels/document.xml.rels" {
                    return Ok(Some(replace_once(
                        bytes,
                        b"</Relationships>",
                        br#"<Relationship Id="rIdEffectsDuplicate" Type="http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects" Target="stylesWithEffects.xml"/></Relationships>"#,
                    )?));
                }
                Ok(Some(bytes.to_vec()))
            },
            &[],
        )?,
        "malformed_third_orphan" => {
            let xml = member_bytes(source)?
                .get("word/stylesWithEffects.xml")
                .cloned()
                .ok_or("main effects source missing")?;
            additions.push(("word/orphanStylesWithEffects.xml", xml));
            rewrite_zip(
                source,
                |name, bytes| {
                    if name == "[Content_Types].xml" {
                        return Ok(Some(replace_once(
                            bytes,
                            b"</Types>",
                            br#"<Override PartName="/word/orphanStylesWithEffects.xml" ContentType="application/vnd.ms-word.stylesWithEffects+xml"/></Types>"#,
                        )?));
                    }
                    Ok(Some(bytes.to_vec()))
                },
                &additions,
            )?
        },
        "malformed_external" => rewrite_zip(
            source,
            |name, bytes| {
                if name == "word/_rels/document.xml.rels" {
                    return Ok(Some(replace_once(
                        bytes,
                        b"Target=\"stylesWithEffects.xml\"",
                        b"Target=\"https://example.invalid/stylesWithEffects.xml\" TargetMode=\"External\"",
                    )?));
                }
                Ok(Some(bytes.to_vec()))
            },
            &[],
        )?,
        "malformed_wrong_content_type" => rewrite_zip(
            source,
            |name, bytes| {
                if name == "[Content_Types].xml" {
                    return Ok(Some(replace_once(
                        bytes,
                        EFFECTS_CONTENT_TYPE.as_bytes(),
                        b"application/xml",
                    )?));
                }
                Ok(Some(bytes.to_vec()))
            },
            &[],
        )?,
        "malformed_outbound" => {
            additions.push((
                "word/_rels/stylesWithEffects.xml.rels",
                br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdOutbound" Type="urn:test:outbound" Target="https://example.invalid/outbound" TargetMode="External"/></Relationships>"#.to_vec(),
            ));
            rewrite_zip(source, |_name, bytes| Ok(Some(bytes.to_vec())), &additions)?
        },
        "malformed_shared_inbound" => {
            let shared = br#"<Relationship Id="rIdSharedEffects" Type="urn:test:shared-inbound" Target="../stylesWithEffects.xml"/>"#;
            rewrite_zip(
                source,
                |name, bytes| {
                    if name == "word/glossary/_rels/document.xml.rels" {
                        return Ok(Some(replace_once(
                            bytes,
                            b"</Relationships>",
                            &[shared.as_slice(), b"</Relationships>"].concat(),
                        )?));
                    }
                    Ok(Some(bytes.to_vec()))
                },
                &[],
            )?
        },
        "malformed_root" => rewrite_zip(
            source,
            |name, bytes| {
                if name == "word/stylesWithEffects.xml" {
                    return Ok(Some(replace_once(bytes, b"<w:styles", b"<w:document")?));
                }
                Ok(Some(bytes.to_vec()))
            },
            &[],
        )?,
        "malformed_namespace" => rewrite_zip(
            source,
            |name, bytes| {
                if name == "word/stylesWithEffects.xml" {
                    return Ok(Some(replace_once(
                        bytes,
                        TRANSITIONAL_W.as_bytes(),
                        b"urn:invalid:wordprocessingml",
                    )?));
                }
                Ok(Some(bytes.to_vec()))
            },
            &[],
        )?,
        _ => return Err(format!("unknown malformed lane: {lane}").into()),
    };
    Ok(output)
}

fn rewrite_zip<F>(source: &[u8], mut transform: F, additions: &[(&str, Vec<u8>)]) -> Result<Vec<u8>>
where
    F: FnMut(&str, &[u8]) -> Result<Option<Vec<u8>>>,
{
    let mut archive = ZipArchive::new(Cursor::new(source))?;
    let mut output = Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(&mut output);
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    let mut existing = BTreeSet::new();
    for index in 0..archive.len() {
        let mut file = archive.by_index(index)?;
        let name = file.name().to_owned();
        existing.insert(name.clone());
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes)?;
        let Some(bytes) = transform(&name, &bytes)? else {
            continue;
        };
        if name.ends_with('/') {
            writer.add_directory(name, options)?;
        } else {
            writer.start_file(name, options)?;
            writer.write_all(&bytes)?;
        }
    }
    for (name, bytes) in additions {
        if existing.contains(*name) {
            return Err(format!("ZIP addition already exists: {name}").into());
        }
        writer.start_file(*name, options)?;
        writer.write_all(bytes)?;
    }
    writer.finish()?;
    Ok(output.into_inner())
}

fn replace_once(source: &[u8], marker: &[u8], replacement: &[u8]) -> Result<Vec<u8>> {
    let offset = source
        .windows(marker.len())
        .position(|window| window == marker)
        .ok_or_else(|| format!("fixture marker {:?} is absent", marker))?;
    let mut output = Vec::with_capacity(
        source
            .len()
            .saturating_sub(marker.len())
            .saturating_add(replacement.len()),
    );
    output.extend_from_slice(&source[..offset]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[offset + marker.len()..]);
    Ok(output)
}

fn remove_member(source: &[u8], name: &str) -> Result<Vec<u8>> {
    rewrite_zip(
        source,
        |current, bytes| {
            if current == name {
                Ok(None)
            } else {
                Ok(Some(bytes.to_vec()))
            }
        },
        &[],
    )
}

#[allow(dead_code)]
fn _pack_uri_display(uri: &PackURI) -> String {
    uri.as_str().to_owned()
}
