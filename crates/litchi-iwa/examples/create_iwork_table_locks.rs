//! Create Pages, Numbers, and Keynote files with locked native tables.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::pages::{PagesDocumentBuilder, PagesEditor};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::table::lock::State as LegacyTableLockState;
use litchi_numbers::table::lock::State as NumbersLockState;
use litchi_numbers::{Package as NumbersPackage, SheetSelector, TableSelector};
use tempfile::NamedTempFile;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_table_locks <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;

    let numbers_path = output.join("table-lock.numbers");
    let numbers = NumbersDocumentBuilder::new()
        .table_name("Locked Table")
        .table_dimensions(4, 3)
        .build()?;
    let numbers_table = numbers.tables()?.remove(0);
    let focused_numbers = NumbersPackage::from_bytes(&numbers.to_bytes()?)?;
    let mut lock = focused_numbers.edit_table_lock(
        SheetSelector::index(0),
        TableSelector::name(&numbers_table.name),
    )?;
    lock.lock();
    let locked_numbers = lock.commit()?;
    write_new(&numbers_path, locked_numbers.package())?;
    assert_eq!(
        NumbersPackage::open(&numbers_path)?.table_lock(
            SheetSelector::index(0),
            TableSelector::name(&numbers_table.name)
        )?,
        NumbersLockState::Locked
    );
    let pages_path = output.join("table-lock.pages");
    let mut pages = PagesDocumentBuilder::new()
        .body_text("Created from scratch with litchi-iwa.\n")
        .body_table("Locked Table", 4, 3)
        .build()?;
    let pages_table = pages.tables()?.remove(0);
    pages.set_body_table_lock_state(pages_table.model_object_id, LegacyTableLockState::Locked)?;
    pages.save(&pages_path)?;
    assert_eq!(
        PagesEditor::open(&pages_path)?.body_table_lock_state(pages_table.model_object_id)?,
        LegacyTableLockState::Locked
    );

    let keynote_path = output.join("table-lock.key");
    let mut keynote = KeynoteDocumentBuilder::new()
        .title("Locked native table")
        .build()?;
    let keynote_table = keynote.add_slide_table(
        0,
        "Locked Table",
        4,
        3,
        DrawablePoint { x: 320.0, y: 360.0 },
        DrawableSize {
            width: 1_280.0,
            height: 480.0,
        },
    )?;
    keynote.set_slide_table_lock_state(
        0,
        keynote_table.drawable_object_id,
        LegacyTableLockState::Locked,
    )?;
    keynote.save(&keynote_path)?;
    assert_eq!(
        KeynoteEditor::open(&keynote_path)?
            .slide_table_lock_state(0, keynote_table.drawable_object_id)?,
        LegacyTableLockState::Locked
    );
    Ok(())
}

fn write_new(path: &Path, package: &NumbersPackage) -> Result<(), Box<dyn std::error::Error>> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let mut temporary = NamedTempFile::new_in(parent)?;
    package.write_to(temporary.as_file_mut())?;
    temporary.as_file().sync_all()?;
    temporary
        .persist_noclobber(path)
        .map_err(|error| error.error)?;
    Ok(())
}
