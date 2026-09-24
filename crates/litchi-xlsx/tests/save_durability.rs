//! Change 0761: every opt-in save-durability level publishes exactly the bytes
//! of the default `save`, through the same atomic replacement.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use std::path::{Path, PathBuf};

use litchi_core::Durability;
use litchi_xlsx::{Package, Workbook};

const LEVELS: [Durability; 3] = [Durability::Full, Durability::FileOnly, Durability::NoSync];

fn seeded(directory: &Path, name: &str, durability: Durability) -> PathBuf {
    let destination = directory.join(format!("{name}-{}.xlsx", durability.as_str()));
    std::fs::write(&destination, b"old destination").unwrap();
    destination
}

#[test]
fn every_level_publishes_the_default_save_bytes() {
    let directory = tempfile::tempdir().unwrap();
    let workbook = Workbook::create().unwrap();
    let reference = directory.path().join("reference.xlsx");
    workbook.save(&reference).unwrap();
    let expected = std::fs::read(&reference).unwrap();

    let package = Package::open(&reference).unwrap();
    let package_reference = directory.path().join("package-reference.xlsx");
    package.save(&package_reference).unwrap();
    let package_expected = std::fs::read(&package_reference).unwrap();

    for durability in LEVELS {
        let destination = seeded(directory.path(), "workbook", durability);
        workbook
            .save_with_durability(&destination, durability)
            .unwrap();
        assert_eq!(
            std::fs::read(&destination).unwrap(),
            expected,
            "{durability:?}"
        );

        let destination = seeded(directory.path(), "workbook-plain", durability);
        workbook
            .save_plain_with_durability(&destination, durability)
            .unwrap();
        assert_eq!(
            std::fs::read(&destination).unwrap(),
            expected,
            "{durability:?}"
        );

        let destination = seeded(directory.path(), "package", durability);
        package
            .save_with_durability(&destination, durability)
            .unwrap();
        assert_eq!(
            std::fs::read(&destination).unwrap(),
            package_expected,
            "{durability:?}"
        );

        let destination = seeded(directory.path(), "package-plain", durability);
        package
            .save_plain_with_durability(&destination, durability)
            .unwrap();
        assert_eq!(
            std::fs::read(&destination).unwrap(),
            package_expected,
            "{durability:?}"
        );
    }

    Workbook::open(directory.path().join("workbook-no-sync.xlsx")).unwrap();
    // Two references plus four routes at three levels; no temporary remains.
    assert_eq!(std::fs::read_dir(directory.path()).unwrap().count(), 14);
}
