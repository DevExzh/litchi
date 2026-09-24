//! Change 0761: every opt-in save-durability level publishes exactly the bytes
//! of the default `save`, through the same atomic replacement.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_core::Durability;
use litchi_docx::Package;

const LEVELS: [Durability; 3] = [Durability::Full, Durability::FileOnly, Durability::NoSync];

#[test]
fn every_level_publishes_the_default_save_bytes() {
    let directory = tempfile::tempdir().unwrap();
    let mut package = Package::new().unwrap();
    package
        .document_mut()
        .unwrap()
        .add_paragraph_with_text("save durability 0761");

    let reference = directory.path().join("reference.docx");
    package.save(&reference).unwrap();
    let expected = std::fs::read(&reference).unwrap();

    for durability in LEVELS {
        let destination = directory
            .path()
            .join(format!("{}.docx", durability.as_str()));
        std::fs::write(&destination, b"old destination").unwrap();
        package
            .save_with_durability(&destination, durability)
            .unwrap();
        assert_eq!(
            std::fs::read(&destination).unwrap(),
            expected,
            "{durability:?}"
        );

        let plain = directory
            .path()
            .join(format!("plain-{}.docx", durability.as_str()));
        package
            .save_plain_with_durability(&plain, durability)
            .unwrap();
        assert_eq!(std::fs::read(&plain).unwrap(), expected, "{durability:?}");
    }

    let reopened = Package::open(directory.path().join("no-sync.docx")).unwrap();
    assert!(
        reopened
            .document()
            .unwrap()
            .text()
            .unwrap()
            .contains("save durability 0761")
    );
    // Only the published files remain: no level leaves a temporary behind.
    assert_eq!(std::fs::read_dir(directory.path()).unwrap().count(), 7);
}
