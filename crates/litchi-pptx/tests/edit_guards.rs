#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_pptx::Package;
use tempfile::NamedTempFile;

#[test]
fn built_and_opened_presentations_support_immutable_edits() {
    let source = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    builder.presentation().unwrap().add_slide().unwrap();
    let created = builder.build().unwrap();
    created.save(source.path()).unwrap();

    let opened = Package::open(source.path()).unwrap();
    assert!(created.edit().unwrap().commit().unwrap().patch().is_noop());
    assert!(opened.edit().unwrap().commit().unwrap().patch().is_noop());
    assert_eq!(opened.presentation().unwrap().slide_count().unwrap(), 1);
}

#[test]
fn minimal_package_is_directly_editable_and_replayable() {
    let package = Package::new().unwrap();
    assert_eq!(package.presentation().unwrap().slide_count().unwrap(), 0);
    let mut original = Vec::new();
    package.write_to(&mut original).unwrap();
    let mut edit = package.edit().unwrap();
    edit.add_slide_master().unwrap();
    let (changed, patch) = edit.commit().unwrap().into_parts();
    assert!(!patch.is_noop());
    assert_eq!(
        package
            .presentation()
            .unwrap()
            .slide_masters()
            .unwrap()
            .len(),
        1
    );
    assert_eq!(
        changed
            .presentation()
            .unwrap()
            .slide_masters()
            .unwrap()
            .len(),
        2
    );
    assert!(
        package
            .apply(&patch)
            .unwrap()
            .into_package()
            .same_content(&changed)
            .unwrap()
    );
    assert!(
        changed
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .same_content(&package)
            .unwrap()
    );
    let mut unchanged = Vec::new();
    package.write_to(&mut unchanged).unwrap();
    assert_eq!(original, unchanged);
    let mut bytes = Vec::new();
    changed.write_to(&mut bytes).unwrap();
    assert!(
        Package::from_reader(std::io::Cursor::new(bytes))
            .unwrap()
            .same_content(&changed)
            .unwrap()
    );
}
