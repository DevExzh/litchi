//! Create Pages, Numbers, and Keynote files with native decimal cell formats.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::{CellValue, NumbersDocumentBuilder};
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa::table_cell_number_format::{
    TableCellDecimalPlaces, TableCellNegativeNumberStyle, TableCellNumberFormat,
    TableCellThousandsSeparator,
};

const ROW: usize = 1;
const COLUMN: usize = 1;
const CELL_VALUE: f64 = -1_234.5;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_table_number_formats <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;
    create_numbers(&output.join("table-number-formats.numbers"))?;
    create_pages(&output.join("table-number-formats.pages"))?;
    create_keynote(&output.join("table-number-formats.key"))?;
    Ok(())
}

fn format() -> Result<TableCellNumberFormat, litchi_iwa::Error> {
    Ok(TableCellNumberFormat::new(
        TableCellDecimalPlaces::fixed(2)?,
        TableCellNegativeNumberStyle::Parentheses,
        TableCellThousandsSeparator::Shown,
    ))
}

fn create_numbers(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Number Formats")
        .table_dimensions(3, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).object_id;
    editor.set_cell(table_id, ROW, COLUMN, CellValue::Number(CELL_VALUE))?;
    editor.set_table_cell_number_format(table_id, ROW, COLUMN, format()?)?;
    editor.save(output)?;
    Ok(())
}

fn create_pages(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Created from scratch with a native decimal table-cell format.\n")
        .body_table("Number Formats", 3, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.set_table_cell(table_id, ROW, COLUMN, CellValue::Number(CELL_VALUE))?;
    editor.set_table_cell_number_format(table_id, ROW, COLUMN, format()?)?;
    editor.save(output)?;
    Ok(())
}

fn create_keynote(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Native table-cell number formats")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Number Formats",
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
        CellValue::Number(CELL_VALUE),
    )?;
    editor.set_slide_table_cell_number_format(0, table.model_object_id, ROW, COLUMN, format()?)?;
    editor.save(output)?;
    Ok(())
}
