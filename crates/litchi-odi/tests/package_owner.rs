//! ADR 0034 W3 evidence for the immutable ODI `Package` owner: atomic save,
//! typed caller-selected limits, deterministic fresh creation and exact no-op
//! publication.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odi::{
    Builder, Error, LimitKind, OpenOptions, Package, SemanticLimitKind, SemanticLimits,
    SourceLimitKind, SourceLimits,
};
use std::path::PathBuf;

/// A fresh, empty scratch directory removed when dropped.
struct Scratch(PathBuf);

impl Scratch {
    fn new(name: &str) -> Self {
        let path = std::env::temp_dir().join(format!(
            "litchi-odi-package-owner-{}-{name}",
            std::process::id()
        ));
        drop(std::fs::remove_dir_all(&path));
        std::fs::create_dir_all(&path).unwrap();
        Self(path)
    }

    fn entries(&self) -> usize {
        std::fs::read_dir(&self.0).unwrap().count()
    }
}

impl Drop for Scratch {
    fn drop(&mut self) {
        drop(std::fs::remove_dir_all(&self.0));
    }
}

fn changed(package: &Package) -> litchi_odi::Commit {
    let mut edit = package.edit();
    edit.set_frame_name(0, Some("Renamed".into())).unwrap();
    edit.commit().unwrap()
}

#[test]
fn fresh_creation_is_deterministic_without_ambient_state() {
    let first = Package::new().unwrap();
    assert_eq!(first.as_bytes(), Package::new().unwrap().as_bytes());
    assert_eq!(first.image().frames().len(), 1);
    assert_eq!(Builder::new().build().unwrap().as_bytes(), first.as_bytes());
}

#[test]
fn borrowed_bytes_round_trip_exactly() {
    let source = Package::new().unwrap();
    let reopened = Package::from_bytes(source.as_bytes()).unwrap();
    assert_eq!(reopened.as_bytes(), source.as_bytes());
    let streamed = Package::from_reader(source.as_bytes()).unwrap();
    assert_eq!(streamed.as_bytes(), source.as_bytes());
    assert_eq!(reopened.to_bytes().unwrap(), source.as_bytes());
    let mut sink = Vec::new();
    reopened.write_to(&mut sink).unwrap();
    assert_eq!(sink, source.as_bytes());
    assert!(reopened.tolerance_report().is_clean());
}

#[test]
fn an_exact_no_op_commit_is_the_source_byte_for_byte() {
    let package = Package::new().unwrap();
    let commit = package.edit().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().as_bytes(), package.as_bytes());
    let (published, patch) = commit.into_parts();
    assert_eq!(
        package.apply(&patch).unwrap().into_package().as_bytes(),
        published.as_bytes()
    );
}

#[test]
fn a_real_edit_changes_bytes_and_its_patch_replays_on_the_source() {
    let package = Package::new().unwrap();
    let source_bytes = package.as_bytes().to_vec();
    let commit = changed(&package);
    assert!(commit.diagnostics().changed());
    assert!(!commit.patch().is_noop());
    assert_ne!(commit.package().as_bytes(), package.as_bytes());
    assert_eq!(package.as_bytes(), source_bytes.as_slice());
    let (published, patch) = commit.into_parts();
    assert_eq!(published.image().frames()[0].name(), Some("Renamed"));
    assert_eq!(
        package.apply(&patch).unwrap().into_package().as_bytes(),
        published.as_bytes()
    );
    assert_eq!(
        published
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .as_bytes(),
        package.as_bytes()
    );
    assert!(published.apply(&patch).is_err());
}

#[test]
fn save_atomically_replaces_the_destination_and_leaves_no_staging_file() {
    let scratch = Scratch::new("atomic");
    let path = scratch.0.join("out.odi");
    std::fs::write(&path, b"previous bytes").unwrap();
    let package = changed(&Package::new().unwrap()).into_package();
    package.save(&path).unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), package.as_bytes());
    assert_eq!(scratch.entries(), 1);
    package
        .save_with_durability(&path, litchi_core::Durability::NoSync)
        .unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), package.as_bytes());
    assert_eq!(Package::open(&path).unwrap().as_bytes(), package.as_bytes());
}

#[cfg(unix)]
#[test]
fn save_through_a_symlink_is_refused_and_the_target_keeps_its_bytes() {
    let scratch = Scratch::new("symlink");
    let target = scratch.0.join("target.odi");
    std::fs::write(&target, b"untouched").unwrap();
    let link = scratch.0.join("link.odi");
    std::os::unix::fs::symlink(&target, &link).unwrap();
    let package = Package::new().unwrap();
    assert!(package.save(&link).is_err());
    assert_eq!(std::fs::read(&target).unwrap(), b"untouched");
    assert_eq!(scratch.entries(), 2);
}

#[test]
fn caller_selected_limits_are_typed_at_every_ingress() {
    let package = Package::new().unwrap();
    let bytes = package.as_bytes();
    let length = u64::try_from(bytes.len()).unwrap();
    let exact = SourceLimits::builder()
        .max_input_bytes(length)
        .unwrap()
        .build()
        .unwrap();
    assert!(
        Package::from_bytes_with_options(bytes, OpenOptions::new().source_limits(exact)).is_ok()
    );

    let tiny = SourceLimits::builder()
        .max_input_bytes(16)
        .unwrap()
        .build()
        .unwrap();
    let options = OpenOptions::new().source_limits(tiny);
    for error in [
        Package::from_bytes_with_options(bytes, options).unwrap_err(),
        Package::from_reader_with_options(bytes, options).unwrap_err(),
    ] {
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::Source(SourceLimitKind::InputBytes),
                maximum: 16,
                ..
            }
        ));
    }

    let content = SemanticLimits::builder()
        .max_content_bytes(16)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        Package::from_bytes_with_options(bytes, OpenOptions::new().semantic_limits(content))
            .unwrap_err(),
        Error::Limit {
            kind: LimitKind::Semantic(SemanticLimitKind::ContentBytes),
            ..
        }
    ));

    let one = SemanticLimits::builder()
        .max_units(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(
        Package::from_bytes_with_options(bytes, OpenOptions::new().semantic_limits(one)).is_ok()
    );
    // ODI admits exactly one normative frame, so the unit ceiling is
    // reached but can never be exceeded by a valid package.
}

#[test]
fn a_detected_prepared_package_is_adopted_without_a_second_open() {
    let package = Package::new().unwrap();
    let prepared =
        litchi_odf_common::detect::prepared(package.as_bytes().to_vec(), SourceLimits::default())
            .unwrap();
    assert_eq!(prepared.format(), litchi_core::FileFormat::Odi);
    assert_eq!(
        Package::from_prepared_package(prepared).unwrap().as_bytes(),
        package.as_bytes()
    );
}
