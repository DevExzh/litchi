//! Create Pages, Numbers, and Keynote files with native hidden table axes.

use std::path::PathBuf;

use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::{PagesDocumentBuilder, PagesEditor};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa::table_hidden_axes::{TableAxisIndex, TableHiddenAxes};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_table_hidden_axes <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;
    let hidden = TableHiddenAxes::new([TableAxisIndex::row(2), TableAxisIndex::column(1)])?;

    let numbers_path = output.join("table-hidden-axes.numbers");
    let mut numbers = NumbersDocumentBuilder::new()
        .table_name("Hidden Axes")
        .table_dimensions(6, 4)
        .build()?;
    let numbers_table = numbers.tables()?.remove(0);
    numbers.set_table_hidden_axes(numbers_table.object_id, &hidden)?;
    numbers.save(&numbers_path)?;
    assert_eq!(
        NumbersEditor::open(&numbers_path)?.table_hidden_axes(numbers_table.object_id)?,
        hidden
    );

    let pages_path = output.join("table-hidden-axes.pages");
    let mut pages = PagesDocumentBuilder::new()
        .body_text("Created from scratch with litchi-iwa.\n")
        .body_table("Hidden Axes", 6, 4)
        .build()?;
    let pages_table = pages.tables()?.remove(0);
    pages.set_body_table_hidden_axes(pages_table.model_object_id, &hidden)?;
    pages.save(&pages_path)?;
    assert_eq!(
        PagesEditor::open(&pages_path)?.body_table_hidden_axes(pages_table.model_object_id)?,
        hidden
    );

    let keynote_path = output.join("table-hidden-axes.key");
    let mut keynote = KeynoteDocumentBuilder::new()
        .title("Native hidden table axes")
        .build()?;
    let keynote_table = keynote.add_slide_table(
        0,
        "Hidden Axes",
        6,
        4,
        DrawablePoint { x: 320.0, y: 300.0 },
        DrawableSize {
            width: 1_280.0,
            height: 600.0,
        },
    )?;
    keynote.set_slide_table_hidden_axes(0, keynote_table.model_object_id, &hidden)?;
    keynote.save(&keynote_path)?;
    assert_eq!(
        KeynoteEditor::open(&keynote_path)?
            .slide_table_hidden_axes(0, keynote_table.model_object_id)?,
        hidden
    );
    Ok(())
}
