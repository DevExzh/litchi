//! Create Pages, Numbers, and Keynote files with native table-cell text layouts.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::{CellValue, NumbersDocumentBuilder};
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa::table_cell_layout::{
    TableCellInset, TableCellInsets, TableCellLayout, TableCellTextWrap, TableCellVerticalAlignment,
};

const ROW: usize = 1;
const COLUMN: usize = 1;
const INSET_POINTS: f32 = 8.0;
const CELL_TEXT: &str = "Wrapped text\nwith an 8 pt inset";

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_table_layouts <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;
    create_numbers(&output.join("table-layouts.numbers"))?;
    create_pages(&output.join("table-layouts.pages"))?;
    create_keynote(&output.join("table-layouts.key"))?;
    Ok(())
}

fn layout(alignment: TableCellVerticalAlignment) -> Result<TableCellLayout, litchi_iwa::Error> {
    Ok(TableCellLayout::default()
        .with_text_wrap(TableCellTextWrap::Wrapped)
        .with_vertical_alignment(alignment)
        .with_insets(TableCellInsets::uniform(TableCellInset::from_points(
            INSET_POINTS,
        )?)))
}

fn create_numbers(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Layouts")
        .table_dimensions(3, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).object_id;
    editor.set_cell(table_id, ROW, COLUMN, CellValue::Text(CELL_TEXT.to_owned()))?;
    editor.set_table_cell_layout(
        table_id,
        ROW,
        COLUMN,
        layout(TableCellVerticalAlignment::Middle)?,
    )?;
    editor.save(output)?;
    Ok(())
}

fn create_pages(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Created from scratch with native table-cell text layout.\n")
        .body_table("Layouts", 3, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.set_table_cell(table_id, ROW, COLUMN, CellValue::Text(CELL_TEXT.to_owned()))?;
    editor.set_table_cell_layout(
        table_id,
        ROW,
        COLUMN,
        layout(TableCellVerticalAlignment::Bottom)?,
    )?;
    editor.save(output)?;
    Ok(())
}

fn create_keynote(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Native table-cell text layouts")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Layouts",
        3,
        3,
        DrawablePoint { x: 320.0, y: 360.0 },
        DrawableSize {
            width: 1_280.0,
            height: 480.0,
        },
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        CellValue::Text(CELL_TEXT.to_owned()),
    )?;
    editor.set_slide_table_cell_layout(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        layout(TableCellVerticalAlignment::Middle)?,
    )?;
    editor.save(output)?;
    Ok(())
}
