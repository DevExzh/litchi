//! ADR 0034 W3 evidence for the immutable OTH `Package` owner: atomic save,
//! typed caller-selected limits, deterministic fresh creation and exact no-op
//! publication.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_core::Position;
use litchi_oth::{
    Builder, Error, LimitKind, OpenOptions, Package, SemanticLimitKind, SemanticLimits,
    SourceLimitKind, SourceLimits, paragraph::Paragraph,
};

fn built(paragraphs: &[&str]) -> Package {
    paragraphs
        .iter()
        .fold(Builder::new(), |builder, text| {
            builder.paragraph(Paragraph::new(*text))
        })
        .build()
        .unwrap()
}

/// A fresh, empty directory under the system temporary directory, removed
/// on drop.
struct ScratchDir(std::path::PathBuf);

impl ScratchDir {
    fn new(label: &str) -> Self {
        let path = std::env::temp_dir().join(format!(
            "litchi-oth-package-owner-{label}-{}",
            std::process::id()
        ));
        let _ = std::fs::remove_dir_all(&path);
        std::fs::create_dir_all(&path).unwrap();
        Self(path)
    }

    fn path(&self) -> &std::path::Path {
        &self.0
    }
}

impl Drop for ScratchDir {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}

#[test]
fn save_atomically_replaces_the_destination_and_leaves_no_staging_file() {
    let directory = ScratchDir::new("save");
    let path = directory.path().join("out.oth");
    std::fs::write(&path, b"previous bytes").unwrap();
    let package = built(&["atomic"]);
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

    let reopened = Package::open(&path).unwrap();
    assert_eq!(reopened.as_bytes(), package.as_bytes());
}

#[cfg(unix)]
#[test]
fn a_save_through_a_symlink_is_refused_and_the_target_keeps_its_bytes() {
    let directory = ScratchDir::new("symlink");
    let target = directory.path().join("target.oth");
    std::fs::write(&target, b"untouched").unwrap();
    let link = directory.path().join("link.oth");
    std::os::unix::fs::symlink(&target, &link).unwrap();
    let package = built(&["refused"]);
    assert!(package.save(&link).is_err());
    assert_eq!(std::fs::read(&target).unwrap(), b"untouched");
    assert_eq!(std::fs::read_dir(directory.path()).unwrap().count(), 2);
}

#[test]
fn from_bytes_round_trips_the_exact_source_bytes() {
    let bytes = built(&["round trip"]).to_bytes().unwrap();
    let package = Package::from_bytes(&bytes).unwrap();
    assert_eq!(package.as_bytes(), bytes.as_slice());
    let streamed = Package::from_reader(bytes.as_slice()).unwrap();
    assert_eq!(streamed.as_bytes(), bytes.as_slice());
    assert_eq!(
        package.document().text_body().unwrap().paragraphs()[0].text(),
        "round trip"
    );
}

#[test]
fn caller_selected_source_limits_are_typed_at_every_ingress() {
    let bytes = built(&["limited"]).to_bytes().unwrap();
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
}

#[test]
fn semantic_limits_bound_the_projection() {
    let one_block = built(&["one paragraph"]).to_bytes().unwrap();
    let one = SemanticLimits::builder()
        .max_units(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(
        Package::from_bytes_with_options(&one_block, OpenOptions::new().semantic_limits(one))
            .is_ok()
    );

    let two_blocks = built(&["first", "second"]).to_bytes().unwrap();
    assert!(matches!(
        Package::from_bytes_with_options(&two_blocks, OpenOptions::new().semantic_limits(one))
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
        Package::from_bytes_with_options(&one_block, OpenOptions::new().semantic_limits(content))
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
        built(&["same input"]).as_bytes(),
        built(&["same input"]).as_bytes()
    );
    assert!(
        Package::new()
            .unwrap()
            .document()
            .text_body()
            .unwrap()
            .paragraphs()
            .is_empty()
    );
}

#[test]
fn an_exact_no_op_commit_is_the_source_byte_for_byte() {
    let bytes = built(&["original"]).to_bytes().unwrap();
    let package = Package::from_bytes(&bytes).unwrap();
    let commit = package.edit().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().as_bytes(), bytes.as_slice());

    let mut edit = package.edit();
    edit.set_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let changed = edit.commit().unwrap();
    assert!(changed.diagnostics().changed());
    assert!(!changed.patch().is_noop());
    assert_ne!(changed.package().as_bytes(), bytes.as_slice());
    assert_eq!(
        package.as_bytes(),
        bytes.as_slice(),
        "the source never changes"
    );
    assert_eq!(
        changed
            .package()
            .document()
            .text_body()
            .unwrap()
            .paragraphs()[0]
            .text(),
        "changed"
    );
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
}

#[test]
fn from_prepared_package_adopts_oth_and_refuses_other_formats() {
    let bytes = built(&["prepared"]).to_bytes().unwrap();
    let prepared = litchi_odf_common::detect::prepared(bytes.clone(), SourceLimits::default())
        .expect("OTH prepares");
    assert_eq!(prepared.format(), litchi_core::detection::FileFormat::Oth);
    let package = Package::from_prepared_package(prepared).unwrap();
    assert_eq!(package.as_bytes(), bytes.as_slice());

    let odt = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/odf/corpus/writer-header-footer.odt"),
    )
    .unwrap();
    let prepared =
        litchi_odf_common::detect::prepared(odt, SourceLimits::default()).expect("ODT prepares");
    assert!(Package::from_prepared_package(prepared).is_err());
}
