#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_pptx::{Error, Package};
use tempfile::NamedTempFile;

#[test]
fn opened_presentation_rejects_legacy_mutable_hydration() {
    let source = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut created = Package::new().unwrap();
    created.presentation_mut().unwrap().add_slide().unwrap();
    created.save(source.path()).unwrap();

    let mut opened = Package::open(source.path()).unwrap();
    assert!(matches!(
        opened.presentation_mut(),
        Err(Error::UnsafeEdit {
            operation: "presentation_mut",
            ..
        })
    ));
    assert_eq!(opened.presentation().unwrap().slide_count().unwrap(), 1);
}
