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
use litchi_ppt::Package;
use litchi_ppt::writer::Writer;

const LEVELS: [Durability; 3] = [Durability::Full, Durability::FileOnly, Durability::NoSync];

/// A private directory removed on drop, including after a failed assertion.
struct ScratchDirectory(PathBuf);

impl std::ops::Deref for ScratchDirectory {
    type Target = PathBuf;

    fn deref(&self) -> &PathBuf {
        &self.0
    }
}

impl AsRef<std::path::Path> for ScratchDirectory {
    fn as_ref(&self) -> &std::path::Path {
        &self.0
    }
}

impl Drop for ScratchDirectory {
    fn drop(&mut self) {
        drop(std::fs::remove_dir_all(&self.0));
    }
}

fn scratch_directory() -> ScratchDirectory {
    static NEXT: AtomicU64 = AtomicU64::new(0);
    let directory = std::env::temp_dir().join(format!(
        "litchi-ppt-0761-{}-{}",
        std::process::id(),
        NEXT.fetch_add(1, Ordering::Relaxed)
    ));
    std::fs::create_dir(&directory).unwrap();
    ScratchDirectory(directory)
}

#[test]
fn every_level_publishes_the_default_save_bytes() {
    let directory = scratch_directory();
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 40, 10, 300, 30, "save durability 0761")
        .unwrap();
    let reference = directory.join("reference.ppt");
    writer.save(&reference).unwrap();
    let expected = std::fs::read(&reference).unwrap();

    for durability in LEVELS {
        let destination = directory.join(format!("{}.ppt", durability.as_str()));
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

    Package::open(directory.join("no-sync.ppt")).unwrap();
    assert_eq!(std::fs::read_dir(&directory).unwrap().count(), 4);
}
