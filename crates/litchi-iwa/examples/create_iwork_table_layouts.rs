//! Create Pages, Numbers, and Keynote files with native table-cell text layouts and alignment.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa::numbers::{CellValue, NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::{PagesDocumentBuilder, PagesEditor};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa::table_cell_layout::{
    TableCellInset, TableCellInsets, TableCellLayout, TableCellTextWrap, TableCellVerticalAlignment,
};
use litchi_iwa::text::{TextAlignment, TextPointSize, TextStyle};

const ROW: usize = 1;
const COLUMN: usize = 1;
const INSET_POINTS: f32 = 8.0;
const NUMBERS_TEXT_POINTS: f32 = 18.0;
const PAGES_TEXT_POINTS: f32 = 17.0;
const KEYNOTE_TEXT_POINTS: f32 = 19.0;
const CELL_TEXT: &str = "Wrapped text\nwith an 8 pt inset";

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let arguments = std::env::args().skip(1).collect::<Vec<_>>();
    let output = PathBuf::from(
        arguments
            .first()
            .ok_or("usage: create_iwork_table_layouts <output-directory> [--verify-only]")?,
    );
    if arguments.get(1).map(String::as_str) != Some("--verify-only") {
        std::fs::create_dir_all(&output)?;
        create_numbers(&output.join("table-layouts.numbers"))?;
        create_pages(&output.join("table-layouts.pages"))?;
        create_keynote(&output.join("table-layouts.key"))?;
    }
    verify(&output)?;
    Ok(())
}

fn verify(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let numbers = NumbersEditor::open(output.join("table-layouts.numbers"))?;
    let numbers_table = numbers.tables()?.remove(0);
    assert_eq!(
        numbers.table_cell_text_alignment(numbers_table.object_id, ROW, COLUMN)?,
        TextAlignment::Center
    );
    assert_eq!(
        numbers.table_cell_text_style(numbers_table.object_id, ROW, COLUMN)?,
        numbers_text_style()?
    );

    let pages = PagesEditor::open(output.join("table-layouts.pages"))?;
    let pages_table = pages.tables()?.remove(0);
    assert_eq!(
        pages.table_cell_text_alignment(pages_table.model_object_id, ROW, COLUMN)?,
        TextAlignment::Right
    );
    assert_eq!(
        pages.table_cell_text_style(pages_table.model_object_id, ROW, COLUMN)?,
        pages_text_style()?
    );

    let keynote = KeynoteEditor::open(output.join("table-layouts.key"))?;
    let keynote_table = keynote.slide_tables(0)?.remove(0);
    assert_eq!(
        keynote.slide_table_cell_text_alignment(0, keynote_table.model_object_id, ROW, COLUMN)?,
        TextAlignment::Justified
    );
    assert_eq!(
        keynote.slide_table_cell_text_style(0, keynote_table.model_object_id, ROW, COLUMN)?,
        keynote_text_style()?
    );
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

fn numbers_text_style() -> Result<TextStyle, litchi_iwa::Error> {
    Ok(TextStyle::new(TextPointSize::from_points(NUMBERS_TEXT_POINTS)?).with_bold(true))
}

fn pages_text_style() -> Result<TextStyle, litchi_iwa::Error> {
    Ok(TextStyle::new(TextPointSize::from_points(PAGES_TEXT_POINTS)?).with_italic(true))
}

fn keynote_text_style() -> Result<TextStyle, litchi_iwa::Error> {
    Ok(
        TextStyle::new(TextPointSize::from_points(KEYNOTE_TEXT_POINTS)?)
            .with_bold(true)
            .with_italic(true),
    )
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
    editor.set_table_cell_text_alignment(table_id, ROW, COLUMN, TextAlignment::Center)?;
    editor.set_table_cell_text_style(table_id, ROW, COLUMN, numbers_text_style()?)?;
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
    editor.set_table_cell_text_alignment(table_id, ROW, COLUMN, TextAlignment::Right)?;
    editor.set_table_cell_text_style(table_id, ROW, COLUMN, pages_text_style()?)?;
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
    editor.set_slide_table_cell_text_alignment(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        TextAlignment::Justified,
    )?;
    editor.set_slide_table_cell_text_style(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        keynote_text_style()?,
    )?;
    editor.save(output)?;
    Ok(())
}
