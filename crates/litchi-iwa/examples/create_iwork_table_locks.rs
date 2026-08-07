//! Create Pages, Numbers, and Keynote files with locked native tables.

use std::path::PathBuf;

use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::{PagesDocumentBuilder, PagesEditor};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::table::lock::State as TableLockState;
use litchi_numbers::TableSelector;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_table_locks <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;

    let numbers_path = output.join("table-lock.numbers");
    let mut numbers = NumbersDocumentBuilder::new()
        .table_name("Locked Table")
        .table_dimensions(4, 3)
        .build()?;
    let numbers_table = numbers.tables()?.remove(0);
    numbers.set_table_lock_state(
        TableSelector::name(&numbers_table.name),
        TableLockState::Locked,
    )?;
    numbers.save(&numbers_path)?;
    assert_eq!(
        NumbersEditor::open(&numbers_path)?
            .table_lock_state(TableSelector::name(&numbers_table.name))?,
        TableLockState::Locked
    );

    let pages_path = output.join("table-lock.pages");
    let mut pages = PagesDocumentBuilder::new()
        .body_text("Created from scratch with litchi-iwa.\n")
        .body_table("Locked Table", 4, 3)
        .build()?;
    let pages_table = pages.tables()?.remove(0);
    pages.set_body_table_lock_state(pages_table.model_object_id, TableLockState::Locked)?;
    pages.save(&pages_path)?;
    assert_eq!(
        PagesEditor::open(&pages_path)?.body_table_lock_state(pages_table.model_object_id)?,
        TableLockState::Locked
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
        TableLockState::Locked,
    )?;
    keynote.save(&keynote_path)?;
    assert_eq!(
        KeynoteEditor::open(&keynote_path)?
            .slide_table_lock_state(0, keynote_table.drawable_object_id)?,
        TableLockState::Locked
    );
    Ok(())
}
