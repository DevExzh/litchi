//! Change 0761: every opt-in save-durability level publishes exactly the bytes
//! of the default `save`, through the same atomic replacement.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_core::Durability;
use litchi_pptx::Package;

const LEVELS: [Durability; 3] = [Durability::Full, Durability::FileOnly, Durability::NoSync];

#[test]
fn every_level_publishes_the_default_save_bytes() {
    let directory = tempfile::tempdir().unwrap();
    let mut package = Package::new().unwrap();
    let reference = directory.path().join("reference.pptx");
    package.save(&reference).unwrap();
    let expected = std::fs::read(&reference).unwrap();

    for durability in LEVELS {
        let destination = directory
            .path()
            .join(format!("{}.pptx", durability.as_str()));
        std::fs::write(&destination, b"old destination").unwrap();
        package
            .save_with_durability(&destination, durability)
            .unwrap();
        assert_eq!(
            std::fs::read(&destination).unwrap(),
            expected,
            "{durability:?}"
        );

        // The explicit plaintext route exists with the encryption feature.
        #[cfg(feature = "encryption")]
        {
            let plain = directory
                .path()
                .join(format!("plain-{}.pptx", durability.as_str()));
            package
                .save_plain_with_durability(&plain, durability)
                .unwrap();
            assert_eq!(std::fs::read(&plain).unwrap(), expected, "{durability:?}");
        }
    }

    Package::open(directory.path().join("no-sync.pptx")).unwrap();
    let published = if cfg!(feature = "encryption") { 7 } else { 4 };
    assert_eq!(
        std::fs::read_dir(directory.path()).unwrap().count(),
        published
    );
}
