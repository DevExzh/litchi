//! ADR 0034 W3 evidence for the immutable Formula `Package` owner: atomic
//! save, typed caller-selected limits, deterministic fresh creation and exact
//! no-op publication.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration-test assertions panic on failure by design"
)]

use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_odf_formula::{
    Builder, Error, LimitKind, NodePath, OpenOptions, Package, SemanticLimitKind, SemanticLimits,
    SourceLimitKind, SourceLimits,
};

const MATHML: &str = r#"<math xmlns="http://www.w3.org/1998/Math/MathML"><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow></math>"#;

/// A unique scratch directory removed on drop (the crate has no `tempfile`
/// dev-dependency).
struct ScratchDirectory(PathBuf);

impl ScratchDirectory {
    fn new(name: &str) -> Self {
        static NEXT: AtomicUsize = AtomicUsize::new(0);
        let path = std::env::temp_dir().join(format!(
            "litchi-odf-formula-{name}-{}-{}",
            std::process::id(),
            NEXT.fetch_add(1, Ordering::Relaxed)
        ));
        drop(std::fs::remove_dir_all(&path));
        std::fs::create_dir_all(&path).unwrap();
        Self(path)
    }

    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for ScratchDirectory {
    fn drop(&mut self) {
        drop(std::fs::remove_dir_all(&self.0));
    }
}

fn fixture() -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/odf/native-resave/source/font-styles.odf"),
    )
    .unwrap()
}

fn built() -> Package {
    Builder::from_mathml(MATHML).build().unwrap()
}

#[test]
fn fresh_creation_is_deterministic_and_reopens() {
    let first = Package::new().unwrap();
    let second = Package::new().unwrap();
    assert_eq!(first.as_bytes(), second.as_bytes());
    assert_eq!(built().as_bytes(), built().as_bytes());
    let reopened = Package::from_bytes(first.as_bytes()).unwrap();
    assert_eq!(reopened.formula().root(), first.formula().root());
    assert!(!first.formula().is_template());
}

#[test]
fn from_bytes_round_trips_the_exact_bytes() {
    let bytes = fixture();
    let package = Package::from_bytes(&bytes).unwrap();
    assert_eq!(package.as_bytes(), bytes.as_slice());
    assert_eq!(package.to_bytes().unwrap(), bytes);
    let mut streamed = Vec::new();
    package.write_to(&mut streamed).unwrap();
    assert_eq!(streamed, bytes);
    let read = Package::from_reader(bytes.as_slice()).unwrap();
    assert_eq!(read.as_bytes(), bytes.as_slice());
}

#[test]
fn an_exact_no_op_commit_is_the_source_byte_for_byte() {
    let bytes = fixture();
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
fn a_real_edit_changes_bytes_and_its_patch_reproduces_them() {
    let package = built();
    let source = package.to_bytes().unwrap();
    let mut edit = package.edit();
    edit.set_text(&NodePath::new([0, 0]), "y").unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert!(!commit.patch().is_noop());
    assert_ne!(commit.package().as_bytes(), source.as_slice());
    assert_eq!(
        package.as_bytes(),
        source.as_slice(),
        "the source never changes"
    );

    let (published, patch) = commit.into_parts();
    assert_eq!(published.formula().root().all_text(), "y+1");
    let replayed = package.apply(&patch).unwrap();
    assert_eq!(replayed.into_package().as_bytes(), published.as_bytes());
    assert_eq!(
        published
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .as_bytes(),
        source.as_slice()
    );
    assert!(published.apply(&patch).is_err());
}

#[test]
fn save_atomically_replaces_the_destination_and_leaves_no_staging_file() {
    let directory = ScratchDirectory::new("save");
    let path = directory.path().join("out.odf");
    std::fs::write(&path, b"previous bytes").unwrap();
    let package = built();
    package.save(&path).unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), package.as_bytes());
    assert_eq!(std::fs::read_dir(directory.path()).unwrap().count(), 1);
    package
        .save_with_durability(&path, litchi_core::Durability::NoSync)
        .unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), package.as_bytes());
    assert_eq!(Package::open(&path).unwrap().as_bytes(), package.as_bytes());
}

#[cfg(unix)]
#[test]
fn a_save_through_a_symlink_is_refused_and_the_target_keeps_its_bytes() {
    let directory = ScratchDirectory::new("symlink");
    let target = directory.path().join("target.odf");
    std::fs::write(&target, b"untouched").unwrap();
    let link = directory.path().join("link.odf");
    std::os::unix::fs::symlink(&target, &link).unwrap();
    assert!(built().save(&link).is_err());
    assert_eq!(std::fs::read(&target).unwrap(), b"untouched");
    assert_eq!(std::fs::read_dir(directory.path()).unwrap().count(), 2);
}

#[test]
fn caller_selected_limits_are_typed_at_every_ingress() {
    let bytes = fixture();
    let length = u64::try_from(bytes.len()).unwrap();
    let exact = SourceLimits::builder()
        .max_input_bytes(length)
        .unwrap()
        .build()
        .unwrap();
    assert!(
        Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(exact)).is_ok()
    );

    let tiny = SourceLimits::builder()
        .max_input_bytes(16)
        .unwrap()
        .build()
        .unwrap();
    let options = OpenOptions::new().source_limits(tiny);
    for error in [
        Package::from_bytes_with_options(&bytes, options).unwrap_err(),
        Package::from_reader_with_options(bytes.as_slice(), options).unwrap_err(),
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
        Package::from_bytes_with_options(&bytes, OpenOptions::new().semantic_limits(content))
            .unwrap_err(),
        Error::Limit {
            kind: LimitKind::Semantic(SemanticLimitKind::ContentBytes),
            maximum: 16,
            ..
        }
    ));
}

#[test]
fn owned_ingress_preserves_the_caller_xml_profile_for_formula_codec() {
    let bytes = fixture();
    let xml_limits = litchi_odf_common::XmlStreamLimits::new(
        64 * 1024 * 1024,
        256,
        131_072,
        512,
        2 * 1024 * 1024,
    )
    .unwrap();
    let package =
        Package::from_bytes_with_options(&bytes, OpenOptions::new().xml_limits(xml_limits))
            .unwrap();
    assert_eq!(package.formula().limits().depth(), 256);
    assert_eq!(package.formula().limits().nodes(), 131_072);
}

#[test]
fn a_prepared_package_is_adopted_and_another_family_is_refused() {
    let bytes = fixture();
    let prepared = litchi_odf_common::detect::prepared(bytes.clone(), SourceLimits::default())
        .expect("prepared Formula");
    let identity = prepared.prepared_index_identity();
    let package = Package::from_prepared_package(prepared).unwrap();
    assert_eq!(package.as_bytes(), bytes.as_slice());
    assert_eq!(package.prepared_index_identity(), identity);

    let text = std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/odf/corpus/writer-header-footer.odt"),
    )
    .unwrap();
    let prepared =
        litchi_odf_common::detect::prepared(text, SourceLimits::default()).expect("prepared text");
    assert!(matches!(
        Package::from_prepared_package(prepared).unwrap_err(),
        Error::Core(litchi_core::Error::UnexpectedFormat { .. })
    ));
}
