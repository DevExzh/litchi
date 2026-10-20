//! ADR 0034 W3 evidence for the immutable ODT `Package` owner: atomic save,
//! typed caller-selected limits, deterministic fresh creation and exact no-op
//! publication.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odt::{
    Builder, Error, LimitKind, OpenOptions, Package, SemanticLimitKind, SemanticLimits,
    SourceLimitKind, SourceLimits,
};

fn corpus() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/odf/corpus/writer-header-footer.odt"),
    )
    .unwrap()
}

fn built(text: &str) -> Package {
    let mut builder = Builder::new();
    builder.add_paragraph(text).unwrap();
    builder.build().unwrap()
}

#[test]
fn save_atomically_replaces_the_destination_and_leaves_no_staging_file() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("out.odt");
    std::fs::write(&path, b"previous bytes").unwrap();
    let package = built("atomic");
    package.save(&path).unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), package.as_bytes());
    assert_eq!(std::fs::read_dir(directory.path()).unwrap().count(), 1);
    package
        .save_with_durability(&path, litchi_core::Durability::NoSync)
        .unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), package.as_bytes());

    let mut streamed = Vec::new();
    package.write_to(&mut streamed).unwrap();
    assert_eq!(streamed, package.to_bytes().unwrap());
}

#[cfg(unix)]
#[test]
fn a_failed_save_leaves_the_destination_unchanged() {
    use std::os::unix::fs::PermissionsExt;

    let directory = tempfile::tempdir().unwrap();
    let target = directory.path().join("target.odt");
    std::fs::write(&target, b"untouched").unwrap();
    let link = directory.path().join("link.odt");
    std::os::unix::fs::symlink(&target, &link).unwrap();
    let package = built("refused");
    assert!(package.save(&link).is_err());
    assert_eq!(std::fs::read(&target).unwrap(), b"untouched");
    assert_eq!(std::fs::read_dir(directory.path()).unwrap().count(), 2);

    // A parent directory that refuses the sibling staging file fails before
    // publication; the existing destination keeps its bytes.
    let locked = directory.path().join("locked");
    std::fs::create_dir(&locked).unwrap();
    let destination = locked.join("out.odt");
    std::fs::write(&destination, b"existing").unwrap();
    std::fs::set_permissions(&locked, std::fs::Permissions::from_mode(0o555)).unwrap();
    // A privileged process ignores directory permissions; only then is the
    // refusal unobservable.
    let enforced = std::fs::File::create(locked.join("probe")).is_err();
    let result = package.save(&destination);
    std::fs::set_permissions(&locked, std::fs::Permissions::from_mode(0o755)).unwrap();
    if enforced {
        assert!(result.is_err());
        assert_eq!(std::fs::read(&destination).unwrap(), b"existing");
        assert_eq!(std::fs::read_dir(&locked).unwrap().count(), 1);
    }
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

    let manifest = SourceLimits::builder()
        .max_manifest_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(manifest))
            .unwrap_err(),
        Error::Limit {
            kind: LimitKind::Source(SourceLimitKind::ManifestBytes),
            ..
        }
    ));
    assert!(matches!(
        SourceLimits::builder().max_members(0).unwrap_err(),
        Error::InvalidLimit {
            kind: LimitKind::Source(SourceLimitKind::Members),
            value: 0,
            ..
        }
    ));
}

#[test]
fn semantic_limits_bound_the_projection() {
    let package = built("one paragraph");
    let bytes = package.to_bytes().unwrap();
    let one = SemanticLimits::builder()
        .max_units(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(
        Package::from_bytes_with_options(&bytes, OpenOptions::new().semantic_limits(one)).is_ok()
    );

    let mut builder = Builder::new();
    builder.add_paragraph("first").unwrap();
    builder.add_paragraph("second").unwrap();
    let two = builder.build().unwrap().to_bytes().unwrap();
    assert!(matches!(
        Package::from_bytes_with_options(&two, OpenOptions::new().semantic_limits(one))
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
        Package::from_bytes_with_options(&bytes, OpenOptions::new().semantic_limits(content))
            .unwrap_err(),
        Error::Limit {
            kind: LimitKind::Semantic(SemanticLimitKind::ContentBytes),
            ..
        }
    ));
}

#[test]
fn fresh_creation_is_deterministic_without_ambient_state() {
    assert_eq!(
        Package::new().unwrap().as_bytes(),
        Package::new().unwrap().as_bytes()
    );
    assert_eq!(
        built("same input").as_bytes(),
        built("same input").as_bytes()
    );
    let meta =
        litchi_odf_common::OwnedPackage::from_bytes(Package::new().unwrap().to_bytes().unwrap())
            .unwrap()
            .get_file("meta.xml")
            .unwrap();
    let meta = String::from_utf8(meta).unwrap();
    assert!(!meta.contains("meta:creation-date"));
    assert!(!meta.contains("<dc:date>"));
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

    let mut edit = package.edit();
    edit.replace_paragraph(litchi_odt::transaction::Position::new(0), "changed")
        .unwrap();
    let changed = edit.commit().unwrap();
    assert!(changed.diagnostics().changed());
    assert!(!changed.patch().is_noop());
    assert_eq!(
        package.as_bytes(),
        bytes.as_slice(),
        "the source never changes"
    );
    assert_eq!(
        changed.package().document().unwrap().paragraphs().unwrap()[0]
            .text()
            .unwrap(),
        "changed"
    );
}
