//! Change 0761: every opt-in save-durability level publishes exactly the bytes
//! of the default `save`, through the same atomic replacement.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use std::io::Cursor;
use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::Durability;
use litchi_xls::Workbook;
use litchi_xls::cell_values::{Reference, Selector};
use litchi_xls::comments::{Snapshot, Value};
use litchi_xls::writer::Writer;

const LEVELS: [Durability; 3] = [Durability::Full, Durability::FileOnly, Durability::NoSync];

fn scratch_directory() -> PathBuf {
    static NEXT: AtomicU64 = AtomicU64::new(0);
    let directory = std::env::temp_dir().join(format!(
        "litchi-xls-0761-{}-{}",
        std::process::id(),
        NEXT.fetch_add(1, Ordering::Relaxed)
    ));
    std::fs::create_dir(&directory).unwrap();
    directory
}

fn writer() -> Writer {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Notes").unwrap();
    writer
        .write_string(sheet, 0, 0, "save durability 0761")
        .unwrap();
    writer.add_comment(sheet, 0, 1, "Author", "note").unwrap();
    writer
}

#[test]
fn every_writer_level_publishes_the_default_save_bytes() {
    let directory = scratch_directory();
    let mut writer = writer();
    let reference = directory.join("reference.xls");
    writer.save(&reference).unwrap();
    let expected = std::fs::read(&reference).unwrap();

    for durability in LEVELS {
        let destination = directory.join(format!("{}.xls", durability.as_str()));
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

    Workbook::new(std::fs::File::open(directory.join("no-sync.xls")).unwrap()).unwrap();
    assert_eq!(std::fs::read_dir(&directory).unwrap().count(), 4);
    std::fs::remove_dir_all(directory).unwrap();
}

#[test]
fn every_source_backed_overlay_level_publishes_the_default_save_bytes() {
    let mut source = Cursor::new(Vec::new());
    writer().write_to(&mut source).unwrap();
    let mut edit = Snapshot::from_bytes(source.into_inner()).unwrap().edit();
    edit.replace(
        Selector::Position(0),
        Reference::new(0, 1).unwrap(),
        Value::new("Author", "edit").unwrap(),
    )
    .unwrap();
    let commit = edit.commit_source_backed().unwrap();
    assert!(!commit.is_noop());

    let directory = scratch_directory();
    let reference = directory.join("reference.xls");
    commit.save(&reference).unwrap();
    let expected = std::fs::read(&reference).unwrap();

    for durability in LEVELS {
        let destination = directory.join(format!("{}.xls", durability.as_str()));
        std::fs::write(&destination, b"old destination").unwrap();
        commit
            .save_with_durability(&destination, durability)
            .unwrap();
        assert_eq!(
            std::fs::read(&destination).unwrap(),
            expected,
            "{durability:?}"
        );
    }

    assert_eq!(std::fs::read_dir(&directory).unwrap().count(), 4);
    std::fs::remove_dir_all(directory).unwrap();
}
