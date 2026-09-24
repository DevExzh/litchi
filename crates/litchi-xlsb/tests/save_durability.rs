//! Change 0761: every opt-in save-durability level publishes exactly the bytes
//! of the default `save`, through the same atomic replacement.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_core::Durability;
use litchi_xlsb::Package;

const LEVELS: [Durability; 3] = [Durability::Full, Durability::FileOnly, Durability::NoSync];

#[test]
fn every_level_publishes_the_default_save_bytes() {
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ooxml/xlsb/sample.xlsb");
    let package = Package::open(&fixture).unwrap();
    let directory = tempfile::tempdir().unwrap();
    let reference = directory.path().join("reference.xlsb");
    package.save(&reference).unwrap();
    let expected = std::fs::read(&reference).unwrap();

    for durability in LEVELS {
        let destination = directory
            .path()
            .join(format!("{}.xlsb", durability.as_str()));
        std::fs::write(&destination, b"old destination").unwrap();
        package
            .save_with_durability(&destination, durability)
            .unwrap();
        assert_eq!(
            std::fs::read(&destination).unwrap(),
            expected,
            "{durability:?}"
        );
    }

    Package::open(directory.path().join("no-sync.xlsb")).unwrap();
    assert_eq!(std::fs::read_dir(directory.path()).unwrap().count(), 4);
}
