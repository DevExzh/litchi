//! Create Numbers, Pages, and Keynote files with native fixed table sections.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::{
    KeynoteDocumentBuilder, KeynoteTableColumnDeletion, KeynoteTableColumnInsertion,
    KeynoteTableHeaderCount, KeynoteTableHeaderSettings, KeynoteTableRowDeletion,
    KeynoteTableRowInsertion,
};
use litchi_iwa::numbers::{
    NumbersDocumentBuilder, NumbersTableHeaderCount, NumbersTableHeaderSettings,
    TableColumnDeletion, TableColumnInsertion, TableRowDeletion, TableRowInsertion,
};
use litchi_iwa::pages::{
    PagesDocumentBuilder, PagesTableColumnDeletion, PagesTableColumnInsertion,
    PagesTableHeaderCount, PagesTableHeaderSettings, PagesTableRowDeletion, PagesTableRowInsertion,
};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_numbers::cell::{Update as TableCellUpdate, Value as CellValue};

const TABLE_ROWS: usize = 4;
const TABLE_COLUMNS: usize = 4;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_table_sections <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;
    create_numbers(
        &output.join("section-insertions.numbers"),
        &output.join("section-deletions.numbers"),
    )?;
    create_pages(
        &output.join("section-insertions.pages"),
        &output.join("section-deletions.pages"),
    )?;
    create_keynote(
        &output.join("section-insertions.key"),
        &output.join("section-deletions.key"),
    )?;
    Ok(())
}

fn create_numbers(insertions: &Path, deletions: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Section CRUD")
        .table_dimensions(TABLE_ROWS, TABLE_COLUMNS)
        .build()?;
    let table_id = editor.tables()?.remove(0).object_id;
    editor.set_table_header_settings(
        table_id,
        NumbersTableHeaderSettings {
            header_rows: Some(NumbersTableHeaderCount::ONE),
            header_columns: Some(NumbersTableHeaderCount::ONE),
            footer_rows: Some(NumbersTableHeaderCount::ONE),
            ..Default::default()
        },
    )?;
    editor.insert_table_row(table_id, TableRowInsertion::header(1))?;
    editor.insert_table_row(table_id, TableRowInsertion::footer(0))?;
    editor.insert_table_column(table_id, TableColumnInsertion::header(1))?;
    editor.set_cells(table_id, section_values())?;
    editor.save(insertions)?;
    editor.remove_table_row(table_id, TableRowDeletion::header(0))?;
    editor.remove_table_row(table_id, TableRowDeletion::footer(1))?;
    editor.remove_table_column(table_id, TableColumnDeletion::header(0))?;
    editor.save(deletions)?;
    Ok(())
}

fn create_pages(insertions: &Path, deletions: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Section-aware table CRUD\n")
        .body_table("Section CRUD", TABLE_ROWS, TABLE_COLUMNS)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.set_table_header_settings(
        table_id,
        PagesTableHeaderSettings {
            header_rows: Some(PagesTableHeaderCount::ONE),
            header_columns: Some(PagesTableHeaderCount::ONE),
            footer_rows: Some(PagesTableHeaderCount::ONE),
            ..Default::default()
        },
    )?;
    editor.insert_table_row(table_id, PagesTableRowInsertion::header(1))?;
    editor.insert_table_row(table_id, PagesTableRowInsertion::footer(0))?;
    editor.insert_table_column(table_id, PagesTableColumnInsertion::header(1))?;
    editor.set_table_cells(table_id, section_values())?;
    editor.save(insertions)?;
    editor.remove_table_row(table_id, PagesTableRowDeletion::header(0))?;
    editor.remove_table_row(table_id, PagesTableRowDeletion::footer(1))?;
    editor.remove_table_column(table_id, PagesTableColumnDeletion::header(0))?;
    editor.save(deletions)?;
    Ok(())
}

fn create_keynote(insertions: &Path, deletions: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Section-aware table CRUD")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Section CRUD",
        TABLE_ROWS,
        TABLE_COLUMNS,
        DrawablePoint { x: 320.0, y: 360.0 },
        DrawableSize {
            width: 1_280.0,
            height: 480.0,
        },
    )?;
    editor.set_slide_table_header_settings(
        0,
        table.model_object_id,
        KeynoteTableHeaderSettings {
            header_rows: Some(KeynoteTableHeaderCount::ONE),
            header_columns: Some(KeynoteTableHeaderCount::ONE),
            footer_rows: Some(KeynoteTableHeaderCount::ONE),
            ..Default::default()
        },
    )?;
    editor.insert_slide_table_row(
        0,
        table.model_object_id,
        KeynoteTableRowInsertion::header(1),
    )?;
    editor.insert_slide_table_row(
        0,
        table.model_object_id,
        KeynoteTableRowInsertion::footer(0),
    )?;
    editor.insert_slide_table_column(
        0,
        table.model_object_id,
        KeynoteTableColumnInsertion::header(1),
    )?;
    editor.set_slide_table_cells(0, table.model_object_id, section_values())?;
    editor.save(insertions)?;
    editor.remove_slide_table_row(0, table.model_object_id, KeynoteTableRowDeletion::header(0))?;
    editor.remove_slide_table_row(0, table.model_object_id, KeynoteTableRowDeletion::footer(1))?;
    editor.remove_slide_table_column(
        0,
        table.model_object_id,
        KeynoteTableColumnDeletion::header(0),
    )?;
    editor.save(deletions)?;
    Ok(())
}

fn section_values() -> Vec<TableCellUpdate> {
    [
        (0, 0, "Header row 1 / column 1"),
        (1, 1, "Header row 2 / column 2"),
        (2, 2, "Body"),
        (4, 3, "Footer row 1"),
        (5, 4, "Footer row 2"),
    ]
    .into_iter()
    .map(|(row, column, value)| {
        TableCellUpdate::new(row, column, CellValue::Text(value.to_owned()))
    })
    .collect()
}
