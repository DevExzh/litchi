//! Change 0761: every opt-in save-durability level publishes exactly the bytes
//! of the default `save`, through the same atomic replacement.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::Durability;
use litchi_doc::Package;
use litchi_doc::writer::Writer;

const LEVELS: [Durability; 3] = [Durability::Full, Durability::FileOnly, Durability::NoSync];

fn scratch_directory() -> PathBuf {
    static NEXT: AtomicU64 = AtomicU64::new(0);
    let directory = std::env::temp_dir().join(format!(
        "litchi-doc-0761-{}-{}",
        std::process::id(),
        NEXT.fetch_add(1, Ordering::Relaxed)
    ));
    std::fs::create_dir(&directory).unwrap();
    directory
}

#[test]
fn every_level_publishes_the_default_save_bytes() {
    let directory = scratch_directory();
    let mut writer = Writer::new();
    writer.add_paragraph("save durability 0761").unwrap();
    let reference = directory.join("reference.doc");
    writer.save(&reference).unwrap();
    let expected = std::fs::read(&reference).unwrap();

    for durability in LEVELS {
        let destination = directory.join(format!("{}.doc", durability.as_str()));
        std::fs::write(&destination, b"old destination").unwrap();
        writer
            .save_with_durability(&destination, durability)
            .unwrap();
        assert_eq!(
            std::fs::read(&destination).unwrap(),
            expected,
            "{durability:?}"
        );
    }

    let mut package = Package::open(directory.join("no-sync.doc")).unwrap();
    assert!(
        package
            .document()
            .unwrap()
            .text()
            .unwrap()
            .contains("save durability 0761")
    );
    assert_eq!(std::fs::read_dir(&directory).unwrap().count(), 4);
    std::fs::remove_dir_all(directory).unwrap();
}
