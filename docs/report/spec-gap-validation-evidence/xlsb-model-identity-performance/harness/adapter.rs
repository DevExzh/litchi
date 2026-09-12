//! Public XLSB identity operations used by the bounded smoke profile.

#![allow(
    dead_code,
    unused_imports,
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::print_stdout,
    reason = "the profile keeps lane metadata and future scale hooks together"
)]

use std::collections::{BTreeMap, BTreeSet};
use std::sync::Arc;
use std::time::Instant;

use litchi_opc::{OpcPackage, PackURI, Part};
use litchi_xlsb::Package;
use litchi_xlsb::data_model::{Error, Limits};
use sha2::{Digest as _, Sha256};

use crate::matrix;

mod fixture {
    include!(concat!(env!("OUT_DIR"), "/data_model_identity.rs"));
    include!("scaled_fixture.rs");

    pub fn complete_no_relationship() -> Package {
        complete_package(false)
    }

    pub fn complete_relationship() -> Package {
        complete_package(true)
    }

    pub fn opaque() -> Package {
        opaque_package()
    }
}

pub const LANES: &[&str] = &[
    "neutral_open_tiny",
    "host_open_tiny",
    "neutral_open_relationship",
    "host_stage_noop_tiny",
    "host_stage_rename_relationship",
    "host_commit_rename_relationship",
    "host_save_reopen_relationship",
    "host_inverse_relationship",
    "host_exact_cap_relationship",
    "host_refusal_opaque",
    "host_refusal_limit",
];

#[derive(Clone)]
pub struct Fixture {
    pub bytes: Arc<[u8]>,
    pub table_count: usize,
    pub relationship_count: usize,
    pub selected_table_id: &'static str,
    pub scale: &'static str,
    pub endpoint_layout: &'static str,
    pub name_profile: &'static str,
    pub recipe_version: u32,
    pub input_sha256: String,
    pub input_fnv1a64: u64,
    pub preservation: PreservationManifest,
    pub semantic: Option<SemanticIdentity>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TableIdentity {
    pub table_id: String,
    pub xml_name: String,
    pub metadata_path: String,
    pub dimension_object_id: String,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RelationshipIdentity {
    pub relationship_id: String,
    pub metadata_path: String,
    pub containing_table: String,
    pub primary_table: String,
    pub primary_column: String,
    pub foreign_column: String,
    pub expected_index_key: String,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TimeGroupingIdentity {
    pub table_name: String,
    pub column_id: String,
    pub column_ids: Vec<String>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SemanticIdentity {
    pub tables: Vec<TableIdentity>,
    pub relationships: Vec<RelationshipIdentity>,
    pub time_groupings: Vec<TimeGroupingIdentity>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SemanticCheck {
    pub table_ids_equal: bool,
    pub table_names_equal: bool,
    pub relationship_ids_equal: bool,
    pub relationship_endpoints_equal: bool,
    pub relationship_paths_equal: bool,
    pub time_grouping_ids_equal: bool,
    pub all_equal: bool,
}

#[derive(Clone, Debug)]
pub struct PreservationManifest {
    pub all_parts: BTreeMap<String, String>,
    pub unchanged_parts: BTreeMap<String, String>,
    pub relationships: BTreeMap<String, String>,
    pub content_types: String,
    pub inner_all: BTreeMap<String, String>,
    pub inner_unchanged: BTreeMap<String, String>,
    pub mutable_inner_paths: BTreeSet<String>,
    pub inner_available: bool,
}

#[derive(Clone, Debug)]
pub struct PreservationCheck {
    pub all_parts_equal: bool,
    pub unchanged_parts_equal: bool,
    pub relationships_equal: bool,
    pub content_types_equal: bool,
    pub inner_all_equal: bool,
    pub inner_unchanged_equal: bool,
    pub inner_member_count: usize,
    pub relationship_owner_count: usize,
    pub content_types_bytes: usize,
}

impl PreservationCheck {
    #[must_use]
    pub fn exact_ok(&self) -> bool {
        self.all_parts_equal
            && self.relationships_equal
            && self.content_types_equal
            && self.inner_all_equal
    }

    #[must_use]
    pub fn changed_ok(&self) -> bool {
        self.unchanged_parts_equal
            && self.relationships_equal
            && self.content_types_equal
            && self.inner_unchanged_equal
    }
}

#[derive(Clone, Copy, Debug, Default)]
pub struct PhaseTimes {
    pub open_ns: Option<u64>,
    pub stage_ns: Option<u64>,
    pub commit_ns: Option<u64>,
    pub save_ns: Option<u64>,
    pub reopen_ns: Option<u64>,
    pub inverse_ns: Option<u64>,
    pub validation_ns: Option<u64>,
}

#[derive(Clone, Debug)]
pub struct ErrorReceipt {
    pub class: String,
    pub message: String,
    pub typed_match: bool,
}

#[derive(Clone, Debug)]
pub struct RunResult {
    pub actual_success: bool,
    pub semantic_ok: bool,
    pub opaque_ok: Option<bool>,
    pub preservation: Option<PreservationCheck>,
    pub source_unchanged: bool,
    pub exact_inverse_ok: Option<bool>,
    pub exact_cap_ok: Option<bool>,
    pub one_under_cap_refused: Option<bool>,
    pub output_bytes: Option<u64>,
    pub staged_bytes: Option<u64>,
    pub candidate_bytes: Option<u64>,
    pub source_bytes: u64,
    pub semantic_observed: Option<SemanticIdentity>,
    pub semantic_check: Option<SemanticCheck>,
    pub phases: PhaseTimes,
    pub error: Option<ErrorReceipt>,
}

#[must_use]
pub fn is_known_lane(lane: &str) -> bool {
    LANES.contains(&lane)
}

#[must_use]
pub fn expected_success(lane: &str) -> bool {
    !matches!(lane, "host_refusal_opaque" | "host_refusal_limit")
}

pub fn fixture_for_lane(lane: &str) -> Result<Fixture, String> {
    if !is_known_lane(lane) {
        return Err(format!("unknown lane: {lane}"));
    }
    let package = if lane.contains("relationship") {
        fixture::complete_relationship()
    } else if lane == "host_refusal_opaque" {
        fixture::opaque()
    } else {
        fixture::complete_no_relationship()
    };
    let semantic = if lane == "host_refusal_opaque" {
        None
    } else {
        Some(capture_semantic_identity(&package)?)
    };
    make_fixture(package, semantic)
}

fn make_fixture(package: Package, semantic: Option<SemanticIdentity>) -> Result<Fixture, String> {
    let bytes = package.to_bytes().map_err(|error| error.to_string())?;
    let snapshot = package.data_model().map_err(|error| error.to_string())?;
    let definition = snapshot
        .definition()
        .ok_or_else(|| String::from("complete smoke fixture has no definition"))?;
    let preservation = capture_manifest(&package, None)?;
    let scale_case = matrix::smoke_case(definition.tables.len(), definition.relationships.len());
    Ok(Fixture {
        input_sha256: sha256_hex(&bytes),
        input_fnv1a64: fnv1a64(&bytes),
        bytes: Arc::from(bytes),
        table_count: definition.tables.len(),
        relationship_count: definition.relationships.len(),
        selected_table_id: "T1",
        scale: scale_case.family,
        endpoint_layout: "selected_table",
        name_profile: "same_length_ascii",
        recipe_version: matrix::RECIPE_VERSION,
        preservation,
        semantic,
    })
}

fn capture_semantic_identity(package: &Package) -> Result<SemanticIdentity, String> {
    let snapshot = package.data_model().map_err(|error| error.to_string())?;
    let definition = snapshot
        .definition()
        .ok_or_else(|| String::from("semantic identity requires a Data Model definition"))?;
    let part = snapshot
        .part()
        .ok_or_else(|| String::from("semantic identity requires a Data Model part"))?;
    let storage = litchi_xldm::inspect_shared(part.bytes()).map_err(|error| error.to_string())?;
    let metadata = litchi_xldm::metadata::inspect(&storage).map_err(|error| error.to_string())?;
    let native = litchi_xldm::native::inspect(&storage, &metadata.native_parse_options())
        .map_err(|error| error.to_string())?;
    let generated = litchi_xldm::generated::inspect_system_generated(&storage)
        .map_err(|error| error.to_string())?;
    let olap =
        litchi_xldm::olap::inspect(&storage, &metadata).map_err(|error| error.to_string())?;
    let closure =
        litchi_xldm::prove_xldm140_closure(&storage, &metadata, &olap, &native, &generated)
            .map_err(|error| error.to_string())?;
    if !closure.is_complete() {
        return Err(String::from(
            "semantic identity closure contains unknown members",
        ));
    }
    let projection = closure.projection();
    let tables = projection
        .tables
        .iter()
        .map(|table| TableIdentity {
            table_id: table.table_id.clone(),
            xml_name: table.xml_name.clone(),
            metadata_path: table.metadata_path.clone(),
            dimension_object_id: table.dimension_object_id.clone(),
        })
        .collect();
    let relationships = projection
        .relationships
        .iter()
        .map(|relationship| RelationshipIdentity {
            relationship_id: relationship.relationship_id.clone(),
            metadata_path: relationship.metadata_path.clone(),
            containing_table: relationship.containing_table.clone(),
            primary_table: relationship.primary_table.clone(),
            primary_column: relationship.primary_column.clone(),
            foreign_column: relationship.foreign_column.clone(),
            expected_index_key: relationship.expected_index_key.clone(),
        })
        .collect();
    let time_groupings = definition
        .time_groupings
        .iter()
        .map(|grouping| TimeGroupingIdentity {
            table_name: grouping.table_name.clone(),
            column_id: grouping.column_id.clone(),
            column_ids: grouping
                .columns
                .iter()
                .map(|column| column.column_id.clone())
                .collect(),
        })
        .collect();
    Ok(SemanticIdentity {
        tables,
        relationships,
        time_groupings,
    })
}

fn expected_semantic_identity(
    fixture: &Fixture,
    renamed_table: Option<&str>,
) -> Option<SemanticIdentity> {
    let expected = fixture.semantic.clone()?;
    renamed_table.map_or(Some(expected.clone()), |name| {
        Some(rename_expected_semantic(
            expected,
            fixture.selected_table_id,
            name,
        ))
    })
}

fn rename_expected_semantic(
    mut expected: SemanticIdentity,
    selected_table_id: &str,
    renamed_table: &str,
) -> SemanticIdentity {
    let Some(selected) = expected
        .tables
        .iter_mut()
        .find(|table| table.table_id == selected_table_id)
    else {
        return expected;
    };
    let old_name = selected.xml_name.clone();
    selected.xml_name = renamed_table.to_owned();
    for relationship in &mut expected.relationships {
        if relationship
            .containing_table
            .eq_ignore_ascii_case(&old_name)
        {
            relationship.containing_table = renamed_table.to_owned();
        }
        if relationship.primary_table.eq_ignore_ascii_case(&old_name) {
            relationship.primary_table = renamed_table.to_owned();
        }
    }
    for grouping in &mut expected.time_groupings {
        if grouping.table_name.eq_ignore_ascii_case(&old_name) {
            grouping.table_name = renamed_table.to_owned();
        }
    }
    expected
}

fn compare_semantic_identity(
    actual: &SemanticIdentity,
    expected: &SemanticIdentity,
) -> SemanticCheck {
    let table_ids_equal = actual
        .tables
        .iter()
        .map(|table| {
            (
                &table.table_id,
                &table.metadata_path,
                &table.dimension_object_id,
            )
        })
        .eq(expected.tables.iter().map(|table| {
            (
                &table.table_id,
                &table.metadata_path,
                &table.dimension_object_id,
            )
        }));
    let table_names_equal = actual
        .tables
        .iter()
        .map(|table| (&table.table_id, &table.xml_name))
        .eq(expected
            .tables
            .iter()
            .map(|table| (&table.table_id, &table.xml_name)));
    let relationship_ids_equal = actual
        .relationships
        .iter()
        .map(|relationship| &relationship.relationship_id)
        .eq(expected
            .relationships
            .iter()
            .map(|relationship| &relationship.relationship_id));
    let relationship_endpoints_equal = actual
        .relationships
        .iter()
        .map(|relationship| {
            (
                &relationship.containing_table,
                &relationship.primary_table,
                &relationship.primary_column,
                &relationship.foreign_column,
                &relationship.expected_index_key,
            )
        })
        .eq(expected.relationships.iter().map(|relationship| {
            (
                &relationship.containing_table,
                &relationship.primary_table,
                &relationship.primary_column,
                &relationship.foreign_column,
                &relationship.expected_index_key,
            )
        }));
    let relationship_paths_equal = actual
        .relationships
        .iter()
        .map(|relationship| &relationship.metadata_path)
        .eq(expected
            .relationships
            .iter()
            .map(|relationship| &relationship.metadata_path));
    let time_grouping_ids_equal = actual.time_groupings == expected.time_groupings;
    let all_equal = table_ids_equal
        && table_names_equal
        && relationship_ids_equal
        && relationship_endpoints_equal
        && relationship_paths_equal
        && time_grouping_ids_equal;
    SemanticCheck {
        table_ids_equal,
        table_names_equal,
        relationship_ids_equal,
        relationship_endpoints_equal,
        relationship_paths_equal,
        time_grouping_ids_equal,
        all_equal,
    }
}

/// Correctness-only result for one deterministic scale/layout/name point.
///
/// This deliberately contains no wall-clock or allocator sample.  The matrix
/// command uses it to prove that each generated closure can be opened,
/// renamed, serialized, reopened, and inverted before any performance run is
/// authorized.
#[derive(Clone, Debug)]
pub struct MatrixResult {
    pub status: String,
    pub expected_success: bool,
    pub family: String,
    pub tables: usize,
    pub relationships: usize,
    pub endpoint_layout: String,
    pub name_profile: String,
    pub source_bytes: usize,
    pub staged_bytes: Option<usize>,
    pub candidate_bytes: Option<usize>,
    pub output_bytes: Option<usize>,
    pub source_sha256: String,
    pub source_unchanged: bool,
    pub no_op_exact: bool,
    pub source_proof_status: String,
    pub source_semantic: Option<SemanticIdentity>,
    pub semantic_observed: Option<SemanticIdentity>,
    pub semantic: Option<SemanticCheck>,
    pub reopened_semantic_observed: Option<SemanticIdentity>,
    pub reopened_semantic: Option<SemanticCheck>,
    pub preservation: Option<MatrixPreservation>,
    pub inverse_exact: bool,
    pub error: Option<ErrorReceipt>,
}

#[derive(Clone, Debug)]
pub struct MatrixPreservation {
    pub source: MatrixMemberManifest,
    pub candidate: MatrixMemberManifest,
    pub mutable_inner_paths: Vec<String>,
}

#[derive(Clone, Debug)]
pub struct MatrixMemberManifest {
    pub parts: BTreeMap<String, String>,
    pub relationships: BTreeMap<String, String>,
    pub content_types: String,
    pub inner_members: BTreeMap<String, String>,
}

fn matrix_name(profile: &str) -> Result<&'static str, String> {
    match profile {
        "same_length_ascii" => Ok("TableX"),
        "shorter" => Ok("T"),
        "longer" => Ok("Table1-renamed-with-more-bytes"),
        "escaped_xml" => Ok("A&B<case>"),
        "unicode" => Ok("表一"),
        _ => Err(format!("unknown name profile: {profile}")),
    }
}

fn matrix_member_manifest(manifest: &PreservationManifest) -> MatrixMemberManifest {
    MatrixMemberManifest {
        parts: manifest.all_parts.clone(),
        relationships: manifest.relationships.clone(),
        content_types: manifest.content_types.clone(),
        inner_members: manifest.inner_all.clone(),
    }
}

fn matrix_preservation(
    source: &PreservationManifest,
    candidate: &PreservationManifest,
    mutable_inner_paths: &BTreeSet<String>,
) -> Result<MatrixPreservation, String> {
    if !source.inner_available || !candidate.inner_available {
        return Err(String::from(
            "complete matrix member preservation requires an inspectable XLDM payload",
        ));
    }
    let source_manifest = matrix_member_manifest(source);
    let candidate_manifest = matrix_member_manifest(candidate);
    if source_manifest
        .parts
        .keys()
        .ne(candidate_manifest.parts.keys())
    {
        return Err(String::from(
            "forward package changed its physical member set",
        ));
    }
    if source_manifest
        .relationships
        .keys()
        .ne(candidate_manifest.relationships.keys())
    {
        return Err(String::from(
            "forward package changed its relationship owner set",
        ));
    }
    if source_manifest
        .inner_members
        .keys()
        .ne(candidate_manifest.inner_members.keys())
    {
        return Err(String::from(
            "forward XLDM rewrite changed its inner member set",
        ));
    }
    let source_inner_paths = source_manifest
        .inner_members
        .keys()
        .cloned()
        .collect::<BTreeSet<_>>();
    if !mutable_inner_paths.is_subset(&source_inner_paths) {
        return Err(String::from(
            "matrix mutable-path scope contains a missing XLDM member",
        ));
    }
    for path in &source_inner_paths {
        if source_manifest.inner_members[path] != candidate_manifest.inner_members[path]
            && !mutable_inner_paths.contains(path)
        {
            return Err(format!(
                "forward XLDM rewrite changed an unadmitted member: {path}"
            ));
        }
    }
    Ok(MatrixPreservation {
        source: source_manifest,
        candidate: candidate_manifest,
        mutable_inner_paths: mutable_inner_paths.iter().cloned().collect(),
    })
}

fn matrix_mutable_inner_paths(source: &SemanticIdentity) -> Result<BTreeSet<String>, String> {
    let mut paths = BTreeSet::from([String::from("BackupLog")]);
    let selected = source
        .tables
        .iter()
        .find(|table| table.table_id == "T1")
        .ok_or_else(|| String::from("matrix source has no selected T1 table"))?;
    paths.insert(selected.metadata_path.clone());
    for relationship in &source.relationships {
        if relationship.containing_table == "T1"
            || relationship.containing_table == "Table1"
            || relationship.primary_table == "T1"
            || relationship.primary_table == "Table1"
        {
            paths.insert(relationship.metadata_path.clone());
        }
    }
    Ok(paths)
}

/// Generate and validate one complete matrix point through the public host
/// transaction API.  The function keeps the source package alive throughout
/// the operation so source immutability and inverse checks exercise the same
/// borrowed-source boundary as the ordinary smoke lanes.
pub fn run_matrix_case(
    case: matrix::ScaleCase,
    endpoint_layout: &str,
    name_profile: &str,
) -> Result<MatrixResult, String> {
    if !matrix::ENDPOINT_LAYOUTS.contains(&endpoint_layout) {
        return Err(format!("unknown endpoint layout: {endpoint_layout}"));
    }
    let renamed_table = matrix_name(name_profile)?;
    let package = fixture::complete_scaled(case.tables, case.relationships, endpoint_layout)?;
    let source_bytes = match package.to_bytes() {
        Ok(bytes) => bytes,
        Err(error) => return Err(error.to_string()),
    };
    let source_sha256 = sha256_hex(&source_bytes);
    let source_identity = match capture_semantic_identity(&package) {
        Ok(identity) => identity,
        Err(error) => {
            if error.contains("graph work") {
                return matrix_graph_limit_refusal(
                    package,
                    case,
                    endpoint_layout,
                    name_profile,
                    source_bytes,
                    source_sha256,
                );
            }
            return Err(error);
        },
    };
    let mutable_inner_paths = matrix_mutable_inner_paths(&source_identity)?;
    let source_preservation = capture_manifest(&package, None)?;
    if source_identity.tables.len() != case.tables
        || source_identity.relationships.len() != case.relationships
    {
        return Err(format!(
            "scaled closure projection mismatch for {}:{}: expected {} tables/{} relationships, got {}/{}",
            case.family,
            endpoint_layout,
            case.tables,
            case.relationships,
            source_identity.tables.len(),
            source_identity.relationships.len()
        ));
    }
    let expected = rename_expected_semantic(source_identity.clone(), "T1", renamed_table);

    let source_snapshot = package.data_model().map_err(|error| error.to_string())?;
    let mut transaction = source_snapshot.edit();
    let changed = transaction
        .rename_table("T1", renamed_table)
        .map_err(|error| error.to_string())?;
    if !changed {
        return Err(String::from("matrix rename unexpectedly became a no-op"));
    }
    let staged_bytes = transaction
        .payload()
        .ok_or_else(|| String::from("matrix transaction lost the model payload"))?
        .len();
    let commit = transaction.commit().map_err(|error| error.to_string())?;
    if !commit.changed() || commit.patch().is_empty() {
        return Err(String::from("matrix rename produced an empty commit"));
    }
    let changed_package = package
        .apply_data_model(&commit)
        .map_err(|error| error.to_string())?;
    let candidate_preservation = capture_manifest(&changed_package, Some(&mutable_inner_paths))?;
    let preservation = matrix_preservation(
        &source_preservation,
        &candidate_preservation,
        &mutable_inner_paths,
    )?;
    let candidate_bytes = commit
        .snapshot()
        .part()
        .ok_or_else(|| String::from("matrix commit lost the model part"))?
        .len();
    let changed_identity = capture_semantic_identity(&changed_package)?;
    let semantic = compare_semantic_identity(&changed_identity, &expected);
    if !semantic.all_equal {
        return Err(format!(
            "matrix semantic check failed for {}:{}:{}: {semantic:?}",
            case.family, endpoint_layout, name_profile
        ));
    }
    let output = changed_package
        .to_bytes()
        .map_err(|error| error.to_string())?;
    let reopened = Package::from_bytes(output.clone()).map_err(|error| error.to_string())?;
    let reopened_identity = capture_semantic_identity(&reopened)?;
    let reopened_semantic = compare_semantic_identity(&reopened_identity, &expected);
    if !reopened_semantic.all_equal {
        return Err(format!(
            "matrix readback semantic check failed for {}:{}:{}: {reopened_semantic:?}",
            case.family, endpoint_layout, name_profile
        ));
    }

    // The original package is detached and source-backed; the transaction
    // must not mutate it while staging or applying the forward patch.
    let source_unchanged = package.to_bytes().map_err(|error| error.to_string())? == source_bytes;
    if !source_unchanged {
        return Err(String::from("matrix source bytes changed during staging"));
    }

    let restored = reopened
        .apply_data_model_patch(&commit.patch().inverse())
        .map_err(|error| error.to_string())?;
    let restored_bytes = restored.to_bytes().map_err(|error| error.to_string())?;
    let inverse_exact = restored_bytes == source_bytes;
    if !inverse_exact {
        return Err(format!(
            "matrix inverse was not byte-exact for {}:{}:{}",
            case.family, endpoint_layout, name_profile
        ));
    }

    let mut no_op = source_snapshot.edit();
    let no_op_changed = no_op
        .rename_table("T1", "Table1")
        .map_err(|error| error.to_string())?;
    if no_op_changed {
        return Err(String::from("same-name matrix rename was not a no-op"));
    }
    let no_op_commit = no_op.commit().map_err(|error| error.to_string())?;
    let no_op_package = package
        .apply_data_model(&no_op_commit)
        .map_err(|error| error.to_string())?;
    let no_op_exact = no_op_package
        .to_bytes()
        .map_err(|error| error.to_string())?
        == source_bytes;
    if !no_op_exact {
        return Err(String::from("same-name matrix rename changed source bytes"));
    }

    Ok(MatrixResult {
        status: String::from("passed"),
        expected_success: true,
        family: case.family.to_owned(),
        tables: case.tables,
        relationships: case.relationships,
        endpoint_layout: endpoint_layout.to_owned(),
        name_profile: name_profile.to_owned(),
        source_bytes: source_bytes.len(),
        staged_bytes: Some(staged_bytes),
        candidate_bytes: Some(candidate_bytes),
        output_bytes: Some(output.len()),
        source_sha256,
        source_unchanged,
        no_op_exact,
        source_proof_status: String::from("complete_xldm140"),
        source_semantic: Some(source_identity),
        semantic_observed: Some(changed_identity),
        semantic: Some(semantic),
        reopened_semantic_observed: Some(reopened_identity),
        reopened_semantic: Some(reopened_semantic),
        preservation: Some(preservation),
        inverse_exact,
        error: None,
    })
}

fn matrix_graph_limit_refusal(
    package: Package,
    case: matrix::ScaleCase,
    endpoint_layout: &str,
    name_profile: &str,
    source_bytes: Vec<u8>,
    source_sha256: String,
) -> Result<MatrixResult, String> {
    let source_snapshot = package.data_model().map_err(|error| error.to_string())?;
    let mut transaction = source_snapshot.edit();
    let error = transaction
        .rename_table("T1", matrix_name(name_profile)?)
        .err()
        .ok_or_else(|| String::from("matrix graph-limit control unexpectedly accepted"))?;
    let typed_match = matches!(
        &error,
        Error::LimitExceeded { resource, .. } if *resource == "graph work"
    );
    if !typed_match {
        return Err(format!(
            "matrix graph-limit control returned an unexpected error: {error}"
        ));
    }
    let source_unchanged = package.to_bytes().map_err(|value| value.to_string())? == source_bytes;
    if !source_unchanged {
        return Err(String::from(
            "matrix graph-limit refusal changed source bytes",
        ));
    }
    Ok(MatrixResult {
        status: String::from("expected_refusal"),
        expected_success: false,
        family: case.family.to_owned(),
        tables: case.tables,
        relationships: case.relationships,
        endpoint_layout: endpoint_layout.to_owned(),
        name_profile: name_profile.to_owned(),
        source_bytes: source_bytes.len(),
        staged_bytes: None,
        candidate_bytes: None,
        output_bytes: None,
        source_sha256,
        source_unchanged,
        no_op_exact: false,
        source_proof_status: String::from("unavailable_graph_work_limit"),
        source_semantic: None,
        semantic_observed: None,
        semantic: None,
        reopened_semantic_observed: None,
        reopened_semantic: None,
        preservation: None,
        inverse_exact: false,
        error: Some(ErrorReceipt {
            class: error_class(&error).to_owned(),
            message: error.to_string(),
            typed_match,
        }),
    })
}

pub fn run_once(lane: &str, fixture: &Fixture) -> Result<RunResult, String> {
    let source_bytes = fixture.bytes.len() as u64;
    match lane {
        "neutral_open_tiny" | "neutral_open_relationship" => neutral_open(fixture),
        "host_open_tiny" => host_open(fixture),
        "host_stage_noop_tiny" => stage_noop(fixture),
        "host_stage_rename_relationship" => stage_rename(fixture),
        "host_commit_rename_relationship" => commit_rename(fixture),
        "host_save_reopen_relationship" => save_reopen(fixture),
        "host_inverse_relationship" => inverse(fixture),
        "host_exact_cap_relationship" => exact_cap(fixture),
        "host_refusal_opaque" => refusal_opaque(fixture),
        "host_refusal_limit" => refusal_limit(fixture),
        _ => Err(format!("unknown lane: {lane}")),
    }
    .map(|mut result| {
        result.source_bytes = source_bytes;
        result
    })
}

fn neutral_open(fixture: &Fixture) -> Result<RunResult, String> {
    let started = Instant::now();
    let package =
        OpcPackage::from_bytes(fixture.bytes.as_ref()).map_err(|error| error.to_string())?;
    let model = package
        .get_part(&PackURI::new("/xl/model/item.data").map_err(|error| error.to_string())?)
        .map_err(|error| error.to_string())?
        .blob();
    let storage = litchi_xldm::inspect_shared(model).map_err(|error| error.to_string())?;
    let metadata = litchi_xldm::metadata::inspect(&storage).map_err(|error| error.to_string())?;
    let native = litchi_xldm::native::inspect(&storage, &metadata.native_parse_options())
        .map_err(|error| error.to_string())?;
    let generated = litchi_xldm::generated::inspect_system_generated(&storage)
        .map_err(|error| error.to_string())?;
    let olap =
        litchi_xldm::olap::inspect(&storage, &metadata).map_err(|error| error.to_string())?;
    let closure =
        litchi_xldm::prove_xldm140_closure(&storage, &metadata, &olap, &native, &generated)
            .map_err(|error| error.to_string())?;
    let semantic_ok = closure.is_complete()
        && closure.projection().tables.len() == fixture.table_count
        && closure.projection().relationships.len() == fixture.relationship_count;
    Ok(success(
        semantic_ok,
        None,
        None,
        PhaseTimes {
            open_ns: Some(elapsed(started)),
            ..PhaseTimes::default()
        },
    ))
}

fn host_open(fixture: &Fixture) -> Result<RunResult, String> {
    let started = Instant::now();
    let package = Package::from_bytes(fixture.bytes.to_vec()).map_err(|error| error.to_string())?;
    let snapshot = package.data_model().map_err(|error| error.to_string())?;
    let semantic_ok = snapshot.definition().is_some_and(|definition| {
        definition.tables.len() == fixture.table_count
            && definition.relationships.len() == fixture.relationship_count
    });
    let open_ns = elapsed(started);
    let validation_started = Instant::now();
    let preservation = preservation_check(&package, fixture)?;
    let validation_ns = elapsed(validation_started);
    success(
        semantic_ok,
        None,
        None,
        PhaseTimes {
            open_ns: Some(open_ns),
            validation_ns: Some(validation_ns),
            ..PhaseTimes::default()
        },
    )
    .with_preservation(preservation, true)
    .with_semantic(&package, fixture, None)
}

fn stage_noop(fixture: &Fixture) -> Result<RunResult, String> {
    let package = package(fixture)?;
    let snapshot = package.data_model().map_err(|error| error.to_string())?;
    let mut transaction = snapshot.edit();
    let started = Instant::now();
    let changed = transaction
        .rename_table("T1", "OldName")
        .map_err(|error| error.to_string())?;
    let staged_bytes = transaction.payload().map(|payload| payload.len() as u64);
    let stage_ns = elapsed(started);
    let commit_started = Instant::now();
    let commit = transaction.commit().map_err(|error| error.to_string())?;
    let commit_ns = elapsed(commit_started);
    let applied = package
        .apply_data_model(&commit)
        .map_err(|error| error.to_string())?;
    let validation_started = Instant::now();
    let preservation = preservation_check(&applied, fixture)?;
    let validation_ns = elapsed(validation_started);
    let ok = !changed && !commit.changed() && commit.patch().is_empty();
    let result = success(
        ok,
        None,
        None,
        PhaseTimes {
            stage_ns: Some(stage_ns),
            commit_ns: Some(commit_ns),
            validation_ns: Some(validation_ns),
            ..PhaseTimes::default()
        },
    )
    .with_staged_bytes(staged_bytes)
    .with_candidate_bytes(model_bytes(commit.snapshot()))
    .with_preservation(preservation, true);
    result.with_semantic(&applied, fixture, None)
}

fn stage_rename(fixture: &Fixture) -> Result<RunResult, String> {
    let package = package(fixture)?;
    let snapshot = package.data_model().map_err(|error| error.to_string())?;
    let mut transaction = snapshot.edit();
    let started = Instant::now();
    let changed = transaction
        .rename_table("T1", "Renamed")
        .map_err(|error| error.to_string())?;
    let staged_bytes = transaction.payload().map(|payload| payload.len() as u64);
    let semantic_ok = changed
        && transaction
            .definition()
            .is_some_and(|definition| definition.tables[0].name == "Renamed");
    let stage_ns = elapsed(started);
    let source_unchanged =
        package.to_bytes().map_err(|error| error.to_string())? == fixture.bytes.as_ref();
    let commit_started = Instant::now();
    let commit = transaction.commit().map_err(|error| error.to_string())?;
    let commit_ns = elapsed(commit_started);
    let changed = package
        .apply_data_model(&commit)
        .map_err(|error| error.to_string())?;
    let validation_started = Instant::now();
    let preservation = preservation_check(&changed, fixture)?;
    let validation_ns = elapsed(validation_started);
    let result = success(
        semantic_ok,
        None,
        None,
        PhaseTimes {
            stage_ns: Some(stage_ns),
            commit_ns: Some(commit_ns),
            validation_ns: Some(validation_ns),
            ..PhaseTimes::default()
        },
    )
    .with_source_unchanged(source_unchanged)
    .with_staged_bytes(staged_bytes)
    .with_candidate_bytes(model_bytes(commit.snapshot()))
    .with_preservation(preservation, false);
    result.with_semantic(&changed, fixture, Some("Renamed"))
}

fn commit_rename(fixture: &Fixture) -> Result<RunResult, String> {
    let package = package(fixture)?;
    let snapshot = package.data_model().map_err(|error| error.to_string())?;
    let mut transaction = snapshot.edit();
    transaction
        .rename_table("T1", "Renamed")
        .map_err(|error| error.to_string())?;
    let started = Instant::now();
    let commit = transaction.commit().map_err(|error| error.to_string())?;
    let commit_ns = elapsed(started);
    let semantic_ok = commit.changed()
        && commit
            .snapshot()
            .definition()
            .is_some_and(|definition| definition.tables[0].name == "Renamed");
    let validation_started = Instant::now();
    let changed = package
        .apply_data_model(&commit)
        .map_err(|error| error.to_string())?;
    let preservation = preservation_check(&changed, fixture)?;
    let validation_ns = elapsed(validation_started);
    let result = success(
        semantic_ok,
        None,
        None,
        PhaseTimes {
            commit_ns: Some(commit_ns),
            validation_ns: Some(validation_ns),
            ..PhaseTimes::default()
        },
    )
    .with_candidate_bytes(model_bytes(commit.snapshot()))
    .with_preservation(preservation, false);
    result.with_semantic(&changed, fixture, Some("Renamed"))
}

fn save_reopen(fixture: &Fixture) -> Result<RunResult, String> {
    let package = package(fixture)?;
    let snapshot = package.data_model().map_err(|error| error.to_string())?;
    let mut transaction = snapshot.edit();
    transaction
        .rename_table("T1", "Renamed")
        .map_err(|error| error.to_string())?;
    let commit = transaction.commit().map_err(|error| error.to_string())?;
    let changed = package
        .apply_data_model(&commit)
        .map_err(|error| error.to_string())?;
    let save_started = Instant::now();
    let bytes = changed.to_bytes().map_err(|error| error.to_string())?;
    let save_ns = elapsed(save_started);
    let reopen_started = Instant::now();
    let reopened = Package::from_bytes(bytes.clone()).map_err(|error| error.to_string())?;
    let reopen_ns = elapsed(reopen_started);
    let semantic_ok = reopened
        .data_model()
        .map_err(|error| error.to_string())?
        .definition()
        .is_some_and(|definition| definition.tables[0].name == "Renamed");
    let validation_started = Instant::now();
    let preservation = preservation_check(&reopened, fixture)?;
    let validation_ns = elapsed(validation_started);
    let source_unchanged =
        package.to_bytes().map_err(|error| error.to_string())? == fixture.bytes.as_ref();
    let result = success(
        semantic_ok,
        None,
        Some(bytes.len() as u64),
        PhaseTimes {
            save_ns: Some(save_ns),
            reopen_ns: Some(reopen_ns),
            validation_ns: Some(validation_ns),
            ..PhaseTimes::default()
        },
    )
    .with_source_unchanged(source_unchanged)
    .with_candidate_bytes(model_bytes(commit.snapshot()))
    .with_preservation(preservation, false);
    result.with_semantic(&reopened, fixture, Some("Renamed"))
}

fn inverse(fixture: &Fixture) -> Result<RunResult, String> {
    let package = package(fixture)?;
    let snapshot = package.data_model().map_err(|error| error.to_string())?;
    let mut transaction = snapshot.edit();
    transaction
        .rename_table("T1", "Renamed")
        .map_err(|error| error.to_string())?;
    let commit = transaction.commit().map_err(|error| error.to_string())?;
    let changed = package
        .apply_data_model(&commit)
        .map_err(|error| error.to_string())?;
    let started = Instant::now();
    let restored = changed
        .apply_data_model_patch(&commit.patch().inverse())
        .map_err(|error| error.to_string())?;
    let inverse_ns = elapsed(started);
    let save_started = Instant::now();
    let bytes = restored.to_bytes().map_err(|error| error.to_string())?;
    let save_ns = elapsed(save_started);
    let exact = bytes.as_slice() == fixture.bytes.as_ref();
    let validation_started = Instant::now();
    let preservation = preservation_check(&restored, fixture)?;
    let validation_ns = elapsed(validation_started);
    let result = success(
        exact,
        None,
        Some(bytes.len() as u64),
        PhaseTimes {
            inverse_ns: Some(inverse_ns),
            save_ns: Some(save_ns),
            validation_ns: Some(validation_ns),
            ..PhaseTimes::default()
        },
    )
    .with_inverse(exact)
    .with_candidate_bytes(model_bytes(commit.snapshot()))
    .with_preservation(preservation, true);
    result.with_semantic(&restored, fixture, None)
}

fn exact_cap(fixture: &Fixture) -> Result<RunResult, String> {
    let package = package(fixture)?;
    let snapshot = package.data_model().map_err(|error| error.to_string())?;
    let mut transaction = snapshot.edit();
    transaction
        .rename_table("T1", "A considerably longer table name")
        .map_err(|error| error.to_string())?;
    let unconstrained = transaction.commit().map_err(|error| error.to_string())?;
    let changed = package
        .apply_data_model(&unconstrained)
        .map_err(|error| error.to_string())?;
    let candidate_workbook_bytes = changed
        .opc_package()
        .get_part(&PackURI::new("/xl/workbook.bin").map_err(|error| error.to_string())?)
        .map_err(|error| error.to_string())?
        .blob()
        .len();

    let mut exact_limits = Limits::DEFAULT;
    exact_limits.max_rewrite_bytes = candidate_workbook_bytes;
    let exact_snapshot = package
        .data_model_with_limits(exact_limits)
        .map_err(|error| error.to_string())?;
    let mut exact_transaction = exact_snapshot.edit();
    exact_transaction
        .rename_table("T1", "A considerably longer table name")
        .map_err(|error| error.to_string())?;
    let started = Instant::now();
    let exact_cap_ok = exact_transaction
        .commit()
        .is_ok_and(|commit| commit.changed());
    let commit_ns = elapsed(started);

    let mut under_limits = Limits::DEFAULT;
    under_limits.max_rewrite_bytes = candidate_workbook_bytes.saturating_sub(1);
    let under_snapshot = package
        .data_model_with_limits(under_limits)
        .map_err(|error| error.to_string())?;
    let mut under_transaction = under_snapshot.edit();
    under_transaction
        .rename_table("T1", "A considerably longer table name")
        .map_err(|error| error.to_string())?;
    let one_under_cap_refused = under_transaction.commit().is_err();
    let source_unchanged =
        package.to_bytes().map_err(|error| error.to_string())? == fixture.bytes.as_ref();
    let validation_started = Instant::now();
    let preservation = preservation_check(&changed, fixture)?;
    let validation_ns = elapsed(validation_started);
    let result = success(
        exact_cap_ok && one_under_cap_refused,
        None,
        None,
        PhaseTimes {
            commit_ns: Some(commit_ns),
            validation_ns: Some(validation_ns),
            ..PhaseTimes::default()
        },
    )
    .with_source_unchanged(source_unchanged)
    .with_caps(exact_cap_ok, one_under_cap_refused)
    .with_candidate_bytes(model_bytes(unconstrained.snapshot()))
    .with_preservation(preservation, false);
    result.with_semantic(&changed, fixture, Some("A considerably longer table name"))
}

fn refusal_opaque(fixture: &Fixture) -> Result<RunResult, String> {
    let package = package(fixture)?;
    let source = package.to_bytes().map_err(|error| error.to_string())?;
    let snapshot = package.data_model().map_err(|error| error.to_string())?;
    let mut transaction = snapshot.edit();
    let started = Instant::now();
    match transaction.rename_table("T1", "Renamed") {
        Ok(_) => Ok(failure(
            "unsupported_feature",
            false,
            "opaque rename accepted",
        )),
        Err(error) => {
            let typed = matches!(
                &error,
                Error::InvalidFormat(message)
                    if message
                        == "XLDM outer identity proof failed: MS-XLDM storage must contain at least three complete 4096-byte pages"
            );
            let unchanged = package.to_bytes().map_err(|value| value.to_string())? == source;
            let validation_started = Instant::now();
            let preservation = preservation_check(&package, fixture)?;
            let validation_ns = elapsed(validation_started);
            Ok(refusal(
                &error,
                typed,
                unchanged,
                PhaseTimes {
                    stage_ns: Some(elapsed(started)),
                    validation_ns: Some(validation_ns),
                    ..PhaseTimes::default()
                },
            )
            .with_preservation(preservation, true))
        },
    }
}

fn refusal_limit(fixture: &Fixture) -> Result<RunResult, String> {
    let package = package(fixture)?;
    let source = package.to_bytes().map_err(|error| error.to_string())?;
    let workbook = package
        .opc_package()
        .get_part(&PackURI::new("/xl/workbook.bin").map_err(|error| error.to_string())?)
        .map_err(|error| error.to_string())?
        .blob()
        .len();
    let mut limits = Limits::DEFAULT;
    limits.max_rewrite_bytes = workbook;
    let snapshot = package
        .data_model_with_limits(limits)
        .map_err(|error| error.to_string())?;
    let mut transaction = snapshot.edit();
    transaction
        .rename_table("T1", "A considerably longer table name")
        .map_err(|error| error.to_string())?;
    let started = Instant::now();
    match transaction.commit() {
        Ok(_) => Ok(failure("limit_exceeded", false, "rewrite limit accepted")),
        Err(error) => {
            let typed = matches!(
                &error,
                Error::InvalidFormat(message)
                    if message == "Data Model rewritten workbook bytes limit exceeded"
            );
            let unchanged = package.to_bytes().map_err(|value| value.to_string())? == source;
            let validation_started = Instant::now();
            let preservation = preservation_check(&package, fixture)?;
            let validation_ns = elapsed(validation_started);
            Ok(refusal(
                &error,
                typed,
                unchanged,
                PhaseTimes {
                    commit_ns: Some(elapsed(started)),
                    validation_ns: Some(validation_ns),
                    ..PhaseTimes::default()
                },
            )
            .with_preservation(preservation, true))
        },
    }
}

fn package(fixture: &Fixture) -> Result<Package, String> {
    Package::from_bytes(fixture.bytes.to_vec()).map_err(|error| error.to_string())
}

fn model_bytes(snapshot: &litchi_xlsb::data_model::Snapshot) -> Option<u64> {
    snapshot.part().map(|part| part.len() as u64)
}

fn success(
    semantic_ok: bool,
    opaque_ok: Option<bool>,
    output_bytes: Option<u64>,
    phases: PhaseTimes,
) -> RunResult {
    RunResult {
        actual_success: true,
        semantic_ok,
        opaque_ok,
        preservation: None,
        source_unchanged: true,
        exact_inverse_ok: None,
        exact_cap_ok: None,
        one_under_cap_refused: None,
        output_bytes,
        staged_bytes: None,
        candidate_bytes: output_bytes,
        source_bytes: 0,
        semantic_observed: None,
        semantic_check: None,
        phases,
        error: None,
    }
}

impl RunResult {
    fn with_source_unchanged(mut self, value: bool) -> Self {
        self.source_unchanged = value;
        self
    }

    fn with_inverse(mut self, value: bool) -> Self {
        self.exact_inverse_ok = Some(value);
        self
    }

    fn with_staged_bytes(mut self, value: Option<u64>) -> Self {
        self.staged_bytes = value;
        self
    }

    fn with_candidate_bytes(mut self, value: Option<u64>) -> Self {
        self.candidate_bytes = value;
        self
    }

    fn with_semantic(
        mut self,
        package: &Package,
        fixture: &Fixture,
        renamed_table: Option<&str>,
    ) -> Result<Self, String> {
        let actual = capture_semantic_identity(package)?;
        let Some(expected) = expected_semantic_identity(fixture, renamed_table) else {
            return Err(String::from(
                "complete semantic fixture identity is unavailable",
            ));
        };
        let check = compare_semantic_identity(&actual, &expected);
        self.semantic_observed = Some(actual);
        self.semantic_check = Some(check);
        Ok(self)
    }

    fn with_caps(mut self, exact: bool, one_under: bool) -> Self {
        self.exact_cap_ok = Some(exact);
        self.one_under_cap_refused = Some(one_under);
        self
    }

    fn with_preservation(mut self, check: PreservationCheck, exact: bool) -> Self {
        self.opaque_ok = Some(if exact {
            check.exact_ok()
        } else {
            check.changed_ok()
        });
        self.preservation = Some(check);
        self
    }
}

fn refusal(
    error: &Error,
    typed_match: bool,
    source_unchanged: bool,
    phases: PhaseTimes,
) -> RunResult {
    RunResult {
        actual_success: false,
        semantic_ok: true,
        opaque_ok: Some(source_unchanged),
        preservation: None,
        source_unchanged,
        exact_inverse_ok: None,
        exact_cap_ok: None,
        one_under_cap_refused: None,
        output_bytes: None,
        staged_bytes: None,
        candidate_bytes: None,
        source_bytes: 0,
        semantic_observed: None,
        semantic_check: None,
        phases,
        error: Some(ErrorReceipt {
            class: error_class(error).to_owned(),
            message: error.to_string(),
            typed_match,
        }),
    }
}

fn failure(class: &str, source_unchanged: bool, message: &str) -> RunResult {
    RunResult {
        actual_success: true,
        semantic_ok: false,
        opaque_ok: None,
        preservation: None,
        source_unchanged,
        exact_inverse_ok: None,
        exact_cap_ok: None,
        one_under_cap_refused: None,
        output_bytes: None,
        staged_bytes: None,
        candidate_bytes: None,
        source_bytes: 0,
        semantic_observed: None,
        semantic_check: None,
        phases: PhaseTimes::default(),
        error: Some(ErrorReceipt {
            class: class.to_owned(),
            message: message.to_owned(),
            typed_match: false,
        }),
    }
}

fn error_class(error: &Error) -> &'static str {
    match error {
        Error::UnsupportedFeature(_) => "unsupported_feature",
        Error::InvalidFormat(_) => "invalid_format",
        Error::LimitExceeded { .. } => "limit_exceeded",
        _ => "other",
    }
}

#[cfg(test)]
mod matrix_fixture_tests {
    use super::fixture;
    use super::matrix_name;

    fn relationship_containing_tables(package: &litchi_xlsb::Package) -> Vec<String> {
        package
            .data_model()
            .expect("fixture data model")
            .definition()
            .expect("fixture definition")
            .relationships
            .iter()
            .map(|relationship| relationship.from_table.clone())
            .collect()
    }

    #[test]
    fn selected_and_distributed_layouts_have_distinct_endpoint_incidence() {
        for (tables, relationships) in [(4, 3), (16, 15)] {
            let selected = fixture::complete_scaled(tables, relationships, "selected_table")
                .expect("selected fixture");
            let distributed = fixture::complete_scaled(tables, relationships, "distributed")
                .expect("distributed fixture");
            let selected_containing = relationship_containing_tables(&selected);
            let distributed_containing = relationship_containing_tables(&distributed);
            assert_eq!(selected_containing.len(), relationships);
            assert_eq!(distributed_containing.len(), relationships);
            assert!(selected_containing.iter().all(|table| table == "Table1"));
            assert!(
                distributed_containing
                    .iter()
                    .collect::<std::collections::BTreeSet<_>>()
                    .len()
                    > 1
            );
            assert_ne!(selected_containing, distributed_containing);
        }
    }

    #[test]
    fn matrix_name_profiles_are_stable_and_distinct() {
        let names = [
            matrix_name("same_length_ascii").expect("same length"),
            matrix_name("shorter").expect("shorter"),
            matrix_name("longer").expect("longer"),
            matrix_name("escaped_xml").expect("escaped"),
            matrix_name("unicode").expect("unicode"),
        ];
        assert_eq!(
            names,
            [
                "TableX",
                "T",
                "Table1-renamed-with-more-bytes",
                "A&B<case>",
                "表一"
            ]
        );
        assert_eq!(
            names
                .iter()
                .collect::<std::collections::BTreeSet<_>>()
                .len(),
            names.len()
        );
    }
}

fn elapsed(started: Instant) -> u64 {
    started.elapsed().as_nanos().try_into().unwrap_or(u64::MAX)
}

fn capture_manifest(
    package: &Package,
    mutable_hint: Option<&BTreeSet<String>>,
) -> Result<PreservationManifest, String> {
    let all_parts = package
        .opc_package()
        .iter_parts()
        .map(|part| (part.partname().as_str().to_owned(), sha256_hex(part.blob())))
        .collect::<BTreeMap<_, _>>();
    let mut unchanged_parts = all_parts.clone();
    unchanged_parts.remove("/xl/workbook.bin");
    unchanged_parts.remove("/xl/model/item.data");

    let root = PackURI::new("/").map_err(|error| error.to_string())?;
    let mut relationships = BTreeMap::new();
    relationships.insert(
        root.as_str().to_owned(),
        sha256_hex(
            package
                .opc_package()
                .source_relationships(&root)
                .map_err(|error| error.to_string())?
                .bytes(),
        ),
    );
    for part in package.opc_package().iter_parts() {
        let owner = part.partname();
        relationships.insert(
            owner.as_str().to_owned(),
            sha256_hex(
                package
                    .opc_package()
                    .source_relationships(owner)
                    .map_err(|error| error.to_string())?
                    .bytes(),
            ),
        );
    }
    let content_types = package
        .opc_package()
        .source_content_types()
        .map_err(|error| error.to_string())?;
    let content_types_hash = sha256_hex(content_types.bytes());

    let mut inner_all = BTreeMap::new();
    let mut mutable_inner_paths = mutable_hint.cloned().unwrap_or_default();
    let mut inner_available = false;
    let model_uri = PackURI::new("/xl/model/item.data").map_err(|error| error.to_string())?;
    if let Ok(model_part) = package.opc_package().get_part(&model_uri) {
        let inner_result = (|| -> Result<(), String> {
            let storage = litchi_xldm::inspect_shared(model_part.blob())
                .map_err(|error| error.to_string())?;
            for (index, entry) in storage.files.iter().enumerate() {
                let bytes = storage
                    .file_stored_bytes(index)
                    .ok_or_else(|| format!("missing XLDM storage bytes for {}", entry.path))?;
                inner_all.insert(entry.path.clone(), sha256_hex(bytes));
            }
            if mutable_hint.is_none() {
                // The storage directory's BackupLog records each member's
                // physical size/offset, so a variable-length identity rewrite
                // legitimately updates this structural index alongside the
                // admitted table and relationship members.
                mutable_inner_paths.insert(String::from("BackupLog"));
                let metadata =
                    litchi_xldm::metadata::inspect(&storage).map_err(|error| error.to_string())?;
                let native =
                    litchi_xldm::native::inspect(&storage, &metadata.native_parse_options())
                        .map_err(|error| error.to_string())?;
                let generated = litchi_xldm::generated::inspect_system_generated(&storage)
                    .map_err(|error| error.to_string())?;
                let olap = litchi_xldm::olap::inspect(&storage, &metadata)
                    .map_err(|error| error.to_string())?;
                let closure = litchi_xldm::prove_xldm140_closure(
                    &storage, &metadata, &olap, &native, &generated,
                )
                .map_err(|error| error.to_string())?;
                for table in &closure.projection().tables {
                    mutable_inner_paths.insert(table.metadata_path.clone());
                }
                for relationship in &closure.projection().relationships {
                    mutable_inner_paths.insert(relationship.metadata_path.clone());
                }
            }
            Ok(())
        })();
        inner_available = inner_result.is_ok();
        if inner_result.is_err() {
            inner_all.clear();
            mutable_inner_paths.clear();
        }
    }
    let mut inner_unchanged = inner_all.clone();
    for path in &mutable_inner_paths {
        inner_unchanged.remove(path);
    }
    Ok(PreservationManifest {
        all_parts,
        unchanged_parts,
        relationships,
        content_types: content_types_hash,
        inner_all,
        inner_unchanged,
        mutable_inner_paths,
        inner_available,
    })
}

fn preservation_check(package: &Package, fixture: &Fixture) -> Result<PreservationCheck, String> {
    let actual = capture_manifest(package, Some(&fixture.preservation.mutable_inner_paths))?;
    Ok(PreservationCheck {
        all_parts_equal: actual.all_parts == fixture.preservation.all_parts,
        unchanged_parts_equal: actual.unchanged_parts == fixture.preservation.unchanged_parts,
        relationships_equal: actual.relationships == fixture.preservation.relationships,
        content_types_equal: actual.content_types == fixture.preservation.content_types,
        inner_all_equal: actual.inner_all == fixture.preservation.inner_all,
        inner_unchanged_equal: actual.inner_unchanged == fixture.preservation.inner_unchanged,
        inner_member_count: actual.inner_all.len(),
        relationship_owner_count: actual.relationships.len(),
        content_types_bytes: package
            .opc_package()
            .source_content_types()
            .map_err(|error| error.to_string())?
            .bytes()
            .len(),
    })
}

fn sha256_hex(bytes: &[u8]) -> String {
    let mut digest = Sha256::new();
    digest.update(bytes);
    let digest = digest.finalize();
    digest.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn fnv1a64(bytes: &[u8]) -> u64 {
    bytes.iter().fold(0xcbf29ce484222325, |hash, byte| {
        (hash ^ u64::from(*byte)).wrapping_mul(0x100000001b3)
    })
}
