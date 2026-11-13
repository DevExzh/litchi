//! ADR 0034 W3 evidence for the immutable ODG `Package` owner: atomic save,
//! typed caller-selected limits, deterministic fresh creation and exact no-op
//! publication.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odg::{
    Builder, Error, LimitKind, OpenOptions, Package, SemanticLimitKind, SemanticLimits,
    SourceLimitKind, SourceLimits, page::Page,
};
use std::path::{Path, PathBuf};

fn corpus() -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/libreoffice-core/desktop/qa/data/3page.odg"),
    )
    .unwrap()
}

/// A fresh, empty per-test directory under the system temporary directory.
struct Scratch(PathBuf);

impl Scratch {
    fn new(name: &str) -> Self {
        let path = std::env::temp_dir().join(format!(
            "litchi-odg-package-owner-{name}-{}",
            std::process::id()
        ));
        let _ = std::fs::remove_dir_all(&path);
        std::fs::create_dir_all(&path).unwrap();
        Self(path)
    }

    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for Scratch {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}

fn with_pages(names: &[&str]) -> Package {
    let mut edit = Package::new().unwrap().edit();
    for name in names {
        edit.add_page(Page::new(*name)).unwrap();
    }
    edit.commit().unwrap().into_package()
}

#[test]
fn fresh_creation_is_deterministic_without_ambient_state() {
    let first = Package::new().unwrap();
    let second = Package::new().unwrap();
    assert_eq!(first.as_bytes(), second.as_bytes());
    assert_eq!(Builder::new().build().unwrap().as_bytes(), first.as_bytes());
    assert!(first.drawing().pages().is_empty());
    assert_eq!(
        with_pages(&["One"]).as_bytes(),
        with_pages(&["One"]).as_bytes()
    );
}

#[test]
fn borrowed_bytes_round_trip_exactly() {
    let bytes = corpus();
    let package = Package::from_bytes(&bytes).unwrap();
    assert_eq!(package.as_bytes(), bytes.as_slice());
    assert_eq!(package.to_bytes().unwrap(), bytes);
    let mut streamed = Vec::new();
    package.write_to(&mut streamed).unwrap();
    assert_eq!(streamed, bytes);
    let from_reader = Package::from_reader(bytes.as_slice()).unwrap();
    assert_eq!(from_reader.as_bytes(), bytes.as_slice());
    let drawing = package.drawing();
    assert_eq!(drawing.pages().len(), package.pages().len());
    assert_eq!(drawing.as_bytes(), bytes.as_slice());
}

#[test]
fn an_exact_no_op_commit_is_the_source_byte_for_byte() {
    let bytes = corpus();
    let package = Package::from_bytes(&bytes).unwrap();
    let commit = package.edit().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().as_bytes(), bytes.as_slice());
    let (published, patch) = commit.into_parts();
    assert_eq!(
        package.apply(&patch).unwrap().into_package().as_bytes(),
        published.as_bytes()
    );
}

#[test]
fn a_real_edit_changes_bytes_and_its_patch_replays_on_the_source() {
    let package = with_pages(&["One"]);
    let source = package.as_bytes().to_vec();
    let mut edit = package.edit();
    edit.set_page_name(0, "Renamed").unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().change_count(), 1);
    assert!(!commit.patch().is_noop());
    assert_ne!(commit.package().as_bytes(), source.as_slice());
    assert_eq!(
        package.as_bytes(),
        source.as_slice(),
        "the source never changes"
    );
    let (published, patch) = commit.into_parts();
    assert_eq!(published.pages()[0].name(), Some("Renamed"));
    let replayed = package.apply(&patch).unwrap();
    assert_eq!(replayed.into_package().as_bytes(), published.as_bytes());
    assert!(published.apply(&patch).is_err(), "stale source is refused");
    let restored = published.apply(&patch.inverse()).unwrap();
    assert_eq!(restored.into_package().as_bytes(), source.as_slice());
}

#[test]
fn save_atomically_replaces_the_destination_and_leaves_no_staging_file() {
    let scratch = Scratch::new("atomic");
    let path = scratch.path().join("out.odg");
    std::fs::write(&path, b"previous bytes").unwrap();
    let package = with_pages(&["Atomic"]);
    package.save(&path).unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), package.as_bytes());
    assert_eq!(std::fs::read_dir(scratch.path()).unwrap().count(), 1);
    package
        .save_with_durability(&path, litchi_core::Durability::NoSync)
        .unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), package.as_bytes());
    assert_eq!(std::fs::read_dir(scratch.path()).unwrap().count(), 1);
    let reopened = Package::open(&path).unwrap();
    assert_eq!(reopened.as_bytes(), package.as_bytes());
}

#[cfg(unix)]
#[test]
fn a_save_through_a_symlink_is_refused_and_the_target_keeps_its_bytes() {
    let scratch = Scratch::new("symlink");
    let target = scratch.path().join("target.odg");
    std::fs::write(&target, b"untouched").unwrap();
    let link = scratch.path().join("link.odg");
    std::os::unix::fs::symlink(&target, &link).unwrap();
    let package = with_pages(&["Refused"]);
    assert!(package.save(&link).is_err());
    assert_eq!(std::fs::read(&target).unwrap(), b"untouched");
    assert_eq!(std::fs::read_dir(scratch.path()).unwrap().count(), 2);
}

#[test]
fn caller_selected_source_limits_are_typed_at_every_ingress() {
    let bytes = corpus();
    let length = u64::try_from(bytes.len()).unwrap();
    let exact = SourceLimits::builder()
        .max_input_bytes(length)
        .unwrap()
        .build()
        .unwrap();
    let opened =
        Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(exact)).unwrap();
    assert_eq!(opened.as_bytes(), bytes.as_slice());

    let below = SourceLimits::builder()
        .max_input_bytes(length - 1)
        .unwrap()
        .build()
        .unwrap();
    let options = OpenOptions::new().source_limits(below);
    for error in [
        Package::from_bytes_with_options(&bytes, options).unwrap_err(),
        Package::from_reader_with_options(bytes.as_slice(), options).unwrap_err(),
        Package::from_template_bytes_with_options(&bytes, options).unwrap_err(),
        Package::from_bytes_with_password_and_options(&bytes, "unused", options).unwrap_err(),
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

    let scratch = Scratch::new("limits");
    let path = scratch.path().join("source.odg");
    std::fs::write(&path, &bytes).unwrap();
    assert!(matches!(
        Package::open_with_options(&path, options).unwrap_err(),
        Error::Limit {
            kind: LimitKind::Source(SourceLimitKind::InputBytes),
            ..
        }
    ));
}

#[test]
fn semantic_limits_bound_the_projection() {
    let one_page = with_pages(&["One"]);
    let two_pages = with_pages(&["One", "Two"]);
    let one = SemanticLimits::builder()
        .max_units(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(
        Package::from_bytes_with_options(
            one_page.as_bytes(),
            OpenOptions::new().semantic_limits(one)
        )
        .is_ok()
    );
    assert!(matches!(
        Package::from_bytes_with_options(
            two_pages.as_bytes(),
            OpenOptions::new().semantic_limits(one)
        )
        .unwrap_err(),
        Error::Limit {
            kind: LimitKind::Semantic(SemanticLimitKind::Units),
            observed: 2,
            maximum: 1,
        }
    ));

    let content = SemanticLimits::builder()
        .max_content_bytes(16)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        Package::from_bytes_with_options(
            one_page.as_bytes(),
            OpenOptions::new().semantic_limits(content)
        )
        .unwrap_err(),
        Error::Limit {
            kind: LimitKind::Semantic(SemanticLimitKind::ContentBytes),
            ..
        }
    ));
}

#[test]
fn templates_keep_their_flavor_through_the_owner() {
    let template = Builder::new().build_template().unwrap();
    assert!(template.is_template());
    assert!(Package::from_bytes(template.as_bytes()).is_err());
    let reopened = Package::from_template_bytes(template.as_bytes()).unwrap();
    assert!(reopened.is_template());
    assert!(reopened.drawing().is_template());
}

#[test]
fn a_prepared_detection_result_is_adopted_without_reindexing() {
    let bytes = corpus();
    let prepared =
        litchi_odf_common::detect::prepared(bytes.clone(), SourceLimits::default()).unwrap();
    let identity = prepared.prepared_index_identity();
    let package = Package::from_prepared_package(prepared).unwrap();
    assert_eq!(package.as_bytes(), bytes.as_slice());
    assert_eq!(package.prepared_index_identity(), identity);

    let template = Builder::new().build_template().unwrap();
    let prepared =
        litchi_odf_common::detect::prepared(template.to_bytes().unwrap(), SourceLimits::default())
            .unwrap();
    assert!(
        Package::from_prepared_package(prepared)
            .unwrap()
            .is_template()
    );

    let text = litchi_odf_common::detect::prepared(
        std::fs::read(
            Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/odf/corpus/writer-header-footer.odt"),
        )
        .unwrap(),
        SourceLimits::default(),
    )
    .unwrap();
    assert!(Package::from_prepared_package(text).is_err());
}
