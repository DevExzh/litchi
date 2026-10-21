//! Change 0761: every opt-in save-durability level publishes exactly the bytes
//! of the default `save`, through the same atomic replacement.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_core::Position;
use std::io::Cursor;
use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::Durability;
use litchi_xls::cell_values::{Reference, Selector};
use litchi_xls::comments::Value;
use litchi_xls::raw::Reader;
use litchi_xls::writer::Builder;

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
        "litchi-xls-0761-{}-{}",
        std::process::id(),
        NEXT.fetch_add(1, Ordering::Relaxed)
    ));
    std::fs::create_dir(&directory).unwrap();
    ScratchDirectory(directory)
}

fn builder() -> Builder {
    let mut builder = Builder::new();
    let sheet = builder.add_worksheet("Notes").unwrap();
    builder
        .write_string(sheet, 0, 0, "save durability 0761")
        .unwrap();
    builder.add_comment(sheet, 0, 1, "Author", "note").unwrap();
    builder
}

#[test]
fn every_writer_level_publishes_the_default_save_bytes() {
    let directory = scratch_directory();
    let mut builder = builder();
    let package = builder.build().unwrap();
    let reference = directory.join("reference.xls");
    package.save(&reference).unwrap();
    let expected = std::fs::read(&reference).unwrap();

    for durability in LEVELS {
        let destination = directory.join(format!("{}.xls", durability.as_str()));
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

    Reader::new(std::fs::File::open(directory.join("no-sync.xls")).unwrap()).unwrap();
    assert_eq!(std::fs::read_dir(&directory).unwrap().count(), 4);
}

#[test]
fn every_source_backed_overlay_level_publishes_the_default_save_bytes() {
    let mut source = Cursor::new(Vec::new());
    builder().build().unwrap().write_to(&mut source).unwrap();
    let mut edit = litchi_xls::SourceBackedWorkbook::from_read_at(std::sync::Arc::new(
        litchi_core::OwnedSource::new(source.into_inner()),
    ))
    .unwrap()
    .edit_comments()
    .unwrap();
    edit.replace(
        Selector::Position(Position::new(0)),
        Reference::new(0, 1).unwrap(),
        Value::new("Author", "edit").unwrap(),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.patch().is_noop());

    let directory = scratch_directory();
    let reference = directory.join("reference.xls");
    commit.package().save(&reference).unwrap();
    let expected = std::fs::read(&reference).unwrap();

    for durability in LEVELS {
        let destination = directory.join(format!("{}.xls", durability.as_str()));
        std::fs::write(&destination, b"old destination").unwrap();
        commit
            .package()
            .save_with_durability(&destination, durability)
            .unwrap();
        assert_eq!(
            std::fs::read(&destination).unwrap(),
            expected,
            "{durability:?}"
        );
    }

    assert_eq!(std::fs::read_dir(&directory).unwrap().count(), 4);
}
