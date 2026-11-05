//! ADR 0034 W3 evidence for the immutable ODB `Package` owner: atomic save,
//! typed caller-selected limits, deterministic fresh creation and exact no-op
//! publication.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odb::{
    Builder, Error, LimitKind, OpenOptions, Package, Query, SemanticLimitKind, SemanticLimits,
    SourceLimitKind, SourceLimits,
};

const FIXTURE: &[u8] =
    include_bytes!("../../../test-data/libreoffice-core/extras/source/database/biblio.odb");

fn scratch(label: &str) -> std::path::PathBuf {
    let directory = std::path::Path::new(env!("CARGO_TARGET_TMPDIR"))
        .join(format!("odb-package-owner-{label}-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&directory);
    std::fs::create_dir_all(&directory).unwrap();
    directory
}

#[test]
fn fresh_creation_is_deterministic_without_ambient_state() {
    let first = Package::new().unwrap();
    assert_eq!(first.as_bytes(), Package::new().unwrap().as_bytes());
    assert_eq!(first.as_bytes(), Builder::new().build().unwrap().as_bytes());
    assert!(first.database().catalog().unwrap().tables().is_empty());
}

#[test]
fn from_bytes_round_trips_the_exact_bytes() {
    let package = Package::from_bytes(FIXTURE).unwrap();
    assert_eq!(package.as_bytes(), FIXTURE);
    assert_eq!(package.to_bytes().unwrap(), FIXTURE);
    let mut streamed = Vec::new();
    package.write_to(&mut streamed).unwrap();
    assert_eq!(streamed, FIXTURE);
    let read = Package::from_reader(FIXTURE).unwrap();
    assert_eq!(read.as_bytes(), FIXTURE);
    let projection = package.database().clone();
    assert_eq!(projection.as_bytes(), FIXTURE);
}

#[test]
fn prepared_detection_is_adopted_and_foreign_formats_are_refused() {
    let prepared =
        litchi_odf_common::detect::prepared(FIXTURE.to_vec(), SourceLimits::default()).unwrap();
    let adopted = Package::from_prepared_package(prepared).unwrap();
    assert_eq!(adopted.as_bytes(), FIXTURE);

    let odt = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/odf/corpus/writer-header-footer.odt"),
    )
    .unwrap();
    let foreign = litchi_odf_common::detect::prepared(odt, SourceLimits::default()).unwrap();
    assert!(Package::from_prepared_package(foreign).is_err());
}

#[test]
fn an_exact_no_op_commit_is_the_source_byte_for_byte() {
    let package = Package::from_bytes(FIXTURE).unwrap();
    let commit = package.edit().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().as_bytes(), FIXTURE);
    let (published, patch) = commit.into_parts();
    assert_eq!(published.as_bytes(), FIXTURE);
    assert_eq!(
        package.apply(&patch).unwrap().into_package().as_bytes(),
        FIXTURE
    );
}

#[test]
fn a_real_edit_changes_bytes_and_replays_through_apply() {
    let package = Package::from_bytes(FIXTURE).unwrap();
    let mut edit = package.edit();
    edit.add_query(Query::new("owner-query", "SELECT 1"))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert!(!commit.patch().is_noop());
    assert_ne!(commit.package().as_bytes(), FIXTURE);
    assert_eq!(package.as_bytes(), FIXTURE, "the source never changes");
    let (published, patch) = commit.into_parts();
    assert!(
        published
            .database()
            .catalog()
            .unwrap()
            .query("owner-query")
            .unwrap()
            .is_some()
    );
    assert_eq!(
        package.apply(&patch).unwrap().into_package().as_bytes(),
        published.as_bytes()
    );
    assert!(published.apply(&patch).is_err());
    assert_eq!(
        published
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .as_bytes(),
        FIXTURE
    );
}

#[test]
fn save_atomically_replaces_the_destination_and_leaves_no_staging_file() {
    let directory = scratch("atomic");
    let path = directory.join("out.odb");
    std::fs::write(&path, b"previous bytes").unwrap();
    let package = Package::new().unwrap();
    package.save(&path).unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), package.as_bytes());
    assert_eq!(std::fs::read_dir(&directory).unwrap().count(), 1);
    package
        .save_with_durability(&path, litchi_core::Durability::NoSync)
        .unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), package.as_bytes());
    assert_eq!(std::fs::read_dir(&directory).unwrap().count(), 1);
    let reopened = Package::open(&path).unwrap();
    assert_eq!(reopened.as_bytes(), package.as_bytes());
    std::fs::remove_dir_all(&directory).unwrap();
}

#[cfg(unix)]
#[test]
fn a_save_through_a_symlink_is_refused_and_the_target_keeps_its_bytes() {
    let directory = scratch("symlink");
    let target = directory.join("target.odb");
    std::fs::write(&target, b"untouched").unwrap();
    let link = directory.join("link.odb");
    std::os::unix::fs::symlink(&target, &link).unwrap();
    let package = Package::new().unwrap();
    assert!(package.save(&link).is_err());
    assert_eq!(std::fs::read(&target).unwrap(), b"untouched");
    assert_eq!(std::fs::read_dir(&directory).unwrap().count(), 2);
    std::fs::remove_dir_all(&directory).unwrap();
}

#[test]
fn caller_selected_limits_are_typed_at_every_ingress() {
    let length = u64::try_from(FIXTURE.len()).unwrap();
    let exact = SourceLimits::builder()
        .max_input_bytes(length)
        .unwrap()
        .build()
        .unwrap();
    assert!(
        Package::from_bytes_with_options(FIXTURE, OpenOptions::new().source_limits(exact)).is_ok()
    );

    let below = SourceLimits::builder()
        .max_input_bytes(length - 1)
        .unwrap()
        .build()
        .unwrap();
    let options = OpenOptions::new().source_limits(below);
    for error in [
        Package::from_bytes_with_options(FIXTURE, options).unwrap_err(),
        Package::from_reader_with_options(FIXTURE, options).unwrap_err(),
    ] {
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::Source(SourceLimitKind::InputBytes),
                maximum,
                ..
            } if maximum == length - 1
        ));
    }

    let content = SemanticLimits::builder()
        .max_content_bytes(16)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        Package::from_bytes_with_options(FIXTURE, OpenOptions::new().semantic_limits(content))
            .unwrap_err(),
        Error::Limit {
            kind: LimitKind::Semantic(SemanticLimitKind::ContentBytes),
            ..
        }
    ));
}
