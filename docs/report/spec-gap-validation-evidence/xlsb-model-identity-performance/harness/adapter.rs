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

mod fixture {
    include!(concat!(env!("OUT_DIR"), "/data_model_identity.rs"));

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
    pub scale: &'static str,
    pub input_sha256: String,
    pub input_fnv1a64: u64,
    pub preservation: PreservationManifest,
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
    pub candidate_bytes: Option<u64>,
    pub source_bytes: u64,
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
    make_fixture(package)
}

fn make_fixture(package: Package) -> Result<Fixture, String> {
    let bytes = package.to_bytes().map_err(|error| error.to_string())?;
    let snapshot = package.data_model().map_err(|error| error.to_string())?;
    let definition = snapshot
        .definition()
        .ok_or_else(|| String::from("complete smoke fixture has no definition"))?;
    let preservation = capture_manifest(&package, None)?;
    Ok(Fixture {
        input_sha256: sha256_hex(&bytes),
        input_fnv1a64: fnv1a64(&bytes),
        bytes: Arc::from(bytes),
        table_count: definition.tables.len(),
        relationship_count: definition.relationships.len(),
        scale: if definition.relationships.is_empty() {
            "tiny"
        } else {
            "relationship"
        },
        preservation,
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
    Ok(success(
        semantic_ok,
        None,
        None,
        PhaseTimes {
            open_ns: Some(open_ns),
            validation_ns: Some(validation_ns),
            ..PhaseTimes::default()
        },
    )
    .with_preservation(preservation, true))
}

fn stage_noop(fixture: &Fixture) -> Result<RunResult, String> {
    let package = package(fixture)?;
    let snapshot = package.data_model().map_err(|error| error.to_string())?;
    let mut transaction = snapshot.edit();
    let started = Instant::now();
    let changed = transaction
        .rename_table("T1", "OldName")
        .map_err(|error| error.to_string())?;
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
    Ok(success(
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
    .with_preservation(preservation, true))
}

fn stage_rename(fixture: &Fixture) -> Result<RunResult, String> {
    let package = package(fixture)?;
    let snapshot = package.data_model().map_err(|error| error.to_string())?;
    let mut transaction = snapshot.edit();
    let started = Instant::now();
    let changed = transaction
        .rename_table("T1", "Renamed")
        .map_err(|error| error.to_string())?;
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
    Ok(success(
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
    .with_preservation(preservation, false))
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
    Ok(success(
        semantic_ok,
        None,
        None,
        PhaseTimes {
            commit_ns: Some(commit_ns),
            validation_ns: Some(validation_ns),
            ..PhaseTimes::default()
        },
    )
    .with_preservation(preservation, false))
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
    Ok(success(
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
    .with_preservation(preservation, false))
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
    Ok(success(
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
    .with_preservation(preservation, true))
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
    Ok(success(
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
    .with_preservation(preservation, false))
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
        candidate_bytes: output_bytes,
        source_bytes: 0,
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
        candidate_bytes: None,
        source_bytes: 0,
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
        candidate_bytes: None,
        source_bytes: 0,
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
