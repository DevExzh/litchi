//! ADR 0034 W3 evidence for the immutable ODM `Package` owner: atomic save,
//! typed caller-selected limits, deterministic fresh creation and exact no-op
//! publication.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odf_common::{
    OpenOptions, SemanticLimitKind, SemanticLimits, SourceLimitKind, SourceLimits,
};
use litchi_odm::{
    Builder, Package,
    transaction::{BodyItemSpec, Error, LimitKind},
};
use std::path::{Path, PathBuf};

/// A unique scratch directory removed on drop (no extra dev-dependency).
struct Scratch(PathBuf);

impl Scratch {
    fn new(name: &str) -> Self {
        let path = std::env::temp_dir().join(format!(
            "litchi-odm-package-owner-{name}-{}",
            std::process::id()
        ));
        drop(std::fs::remove_dir_all(&path));
        std::fs::create_dir_all(&path).unwrap();
        Self(path)
    }

    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for Scratch {
    fn drop(&mut self) {
        drop(std::fs::remove_dir_all(&self.0));
    }
}

fn built(paragraphs: &[&str]) -> Package {
    let mut builder = Builder::new();
    for paragraph in paragraphs {
        builder = builder.body_item(BodyItemSpec::paragraph(*paragraph).unwrap());
    }
    builder.build().unwrap()
}

#[test]
fn fresh_creation_is_deterministic_without_ambient_state() {
    assert_eq!(
        Package::new().unwrap().as_bytes(),
        Package::new().unwrap().as_bytes()
    );
    assert_eq!(
        built(&["same input"]).as_bytes(),
        built(&["same input"]).as_bytes()
    );
    assert!(
        Package::new()
            .unwrap()
            .document()
            .content_xml()
            .contains("<office:text")
    );
}

#[test]
fn borrowed_bytes_round_trip_exactly() {
    let source = built(&["round trip"]);
    let bytes = source.to_bytes().unwrap();
    let reopened = Package::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.as_bytes(), bytes.as_slice());
    let streamed = Package::from_reader(bytes.as_slice()).unwrap();
    assert_eq!(streamed.as_bytes(), bytes.as_slice());
    let mut written = Vec::new();
    reopened.write_to(&mut written).unwrap();
    assert_eq!(written, bytes);
}

#[test]
fn an_exact_no_op_commit_is_the_source_byte_for_byte() {
    let package = built(&["unchanged"]);
    let bytes = package.to_bytes().unwrap();
    let commit = package.edit().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().as_bytes(), bytes.as_slice());

    let mut edit = package.edit();
    edit.add_body_item(BodyItemSpec::paragraph("added").unwrap())
        .unwrap();
    let changed = edit.commit().unwrap();
    assert!(changed.diagnostics().changed());
    assert!(!changed.patch().is_noop());
    assert_ne!(changed.package().as_bytes(), bytes.as_slice());
    assert_eq!(changed.package().document().structure().items().len(), 2);
    let (published, patch) = changed.into_parts();
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
        bytes.as_slice()
    );
    assert_eq!(
        package.as_bytes(),
        bytes.as_slice(),
        "the source never changes"
    );
}

#[test]
fn save_atomically_replaces_the_destination_and_leaves_no_staging_file() {
    let scratch = Scratch::new("save");
    let path = scratch.path().join("out.odm");
    std::fs::write(&path, b"previous bytes").unwrap();
    let package = built(&["atomic"]);
    package.save(&path).unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), package.as_bytes());
    assert_eq!(std::fs::read_dir(scratch.path()).unwrap().count(), 1);
    package
        .save_with_durability(&path, litchi_core::Durability::NoSync)
        .unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), package.as_bytes());
    assert_eq!(Package::open(&path).unwrap().as_bytes(), package.as_bytes());
}

#[cfg(unix)]
#[test]
fn a_save_through_a_symlink_is_refused_and_the_target_keeps_its_bytes() {
    let scratch = Scratch::new("symlink");
    let target = scratch.path().join("target.odm");
    std::fs::write(&target, b"untouched").unwrap();
    let link = scratch.path().join("link.odm");
    std::os::unix::fs::symlink(&target, &link).unwrap();
    let package = built(&["refused"]);
    assert!(package.save(&link).is_err());
    assert_eq!(std::fs::read(&target).unwrap(), b"untouched");
    assert_eq!(std::fs::read_dir(scratch.path()).unwrap().count(), 2);
}

#[test]
fn caller_selected_limits_are_typed() {
    let bytes = built(&["one"]).to_bytes().unwrap();
    let length = u64::try_from(bytes.len()).unwrap();
    let exact = SourceLimits::builder()
        .max_input_bytes(length)
        .unwrap()
        .build()
        .unwrap();
    assert!(
        Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(exact)).is_ok()
    );
    let below = SourceLimits::builder()
        .max_input_bytes(length - 1)
        .unwrap()
        .build()
        .unwrap();
    let options = OpenOptions::new().source_limits(below);
    for error in [
        Package::from_bytes_with_options(&bytes, options).unwrap_err(),
        Package::from_reader_with_options(bytes.as_slice(), options).unwrap_err(),
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
        Package::from_bytes_with_options(&bytes, OpenOptions::new().semantic_limits(content))
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
        Package::from_bytes_with_options(&bytes, OpenOptions::new().semantic_limits(one)).is_ok()
    );
    let two = built(&["first", "second"]).to_bytes().unwrap();
    assert!(matches!(
        Package::from_bytes_with_options(&two, OpenOptions::new().semantic_limits(one))
            .unwrap_err(),
        Error::Limit {
            kind: LimitKind::Semantic(SemanticLimitKind::Units),
            observed: 2,
            maximum: 1,
        }
    ));
}

#[test]
fn a_prepared_package_is_adopted_and_another_family_is_refused() {
    let bytes = built(&["prepared"]).to_bytes().unwrap();
    let prepared =
        litchi_odf_common::detect::prepared(bytes.clone(), SourceLimits::default()).unwrap();
    let adopted = Package::from_prepared_package(prepared).unwrap();
    assert_eq!(adopted.as_bytes(), bytes.as_slice());

    let content = Package::new().unwrap().document().content_xml().to_owned();
    let mut writer = litchi_odf_common::core::PackageWriter::new(
        litchi_odf_common::core::StreamingArchiveLimits::default(),
    );
    writer
        .set_mimetype("application/vnd.oasis.opendocument.text")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    let text = writer.finish_to_bytes().unwrap();
    let prepared = litchi_odf_common::detect::prepared(text, SourceLimits::default()).unwrap();
    assert!(matches!(
        Package::from_prepared_package(prepared).unwrap_err(),
        Error::Core(litchi_core::Error::UnexpectedFormat { .. })
    ));
}
