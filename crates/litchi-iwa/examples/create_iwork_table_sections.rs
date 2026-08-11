//! Create Numbers, Pages, and Keynote files with native fixed table sections.
use litchi_numbers::table::headers::{Count as HeaderCount, Settings as HeaderSettings};

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_numbers::cell::{Update as TableCellUpdate, Value as CellValue};
use litchi_numbers::table::topology::{ColumnDeletion, ColumnInsertion, RowDeletion, RowInsertion};
use litchi_numbers::{Package, SheetSelector, TableSelector};

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
    let table = TableSelector::index(0);
    editor = set_focused_table_headers(
        editor,
        HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            header_columns: Some(HeaderCount::ONE),
            footer_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )?;
    editor.insert_table_row(table, RowInsertion::header(1))?;
    editor.insert_table_row(table, RowInsertion::footer(0))?;
    editor.insert_table_column(table, ColumnInsertion::header(1))?;
    set_numbers_cells(&mut editor, section_values())?;
    editor.save(insertions)?;
    editor.remove_table_row(table, RowDeletion::header(0))?;
    editor.remove_table_row(table, RowDeletion::footer(1))?;
    editor.remove_table_column(table, ColumnDeletion::header(0))?;
    editor.save(deletions)?;
    Ok(())
}

fn set_focused_table_headers(
    editor: NumbersEditor,
    settings: HeaderSettings,
) -> Result<NumbersEditor, Box<dyn std::error::Error>> {
    let package = Package::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_table_headers(SheetSelector::index(0), TableSelector::index(0))?
        .set(settings)
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    Ok(NumbersEditor::from_bytes(&bytes)?)
}

fn create_pages(insertions: &Path, deletions: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Section-aware table CRUD\n")
        .body_table("Section CRUD", TABLE_ROWS, TABLE_COLUMNS)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.set_table_header_settings(
        table_id,
        HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            header_columns: Some(HeaderCount::ONE),
            footer_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )?;
    editor.insert_table_row(table_id, RowInsertion::header(1))?;
    editor.insert_table_row(table_id, RowInsertion::footer(0))?;
    editor.insert_table_column(table_id, ColumnInsertion::header(1))?;
    editor.set_table_cells(table_id, section_values())?;
    editor.save(insertions)?;
    editor.remove_table_row(table_id, RowDeletion::header(0))?;
    editor.remove_table_row(table_id, RowDeletion::footer(1))?;
    editor.remove_table_column(table_id, ColumnDeletion::header(0))?;
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
        HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            header_columns: Some(HeaderCount::ONE),
            footer_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )?;
    editor.insert_slide_table_row(0, table.model_object_id, RowInsertion::header(1))?;
    editor.insert_slide_table_row(0, table.model_object_id, RowInsertion::footer(0))?;
    editor.insert_slide_table_column(0, table.model_object_id, ColumnInsertion::header(1))?;
    editor.set_slide_table_cells(0, table.model_object_id, section_values())?;
    editor.save(insertions)?;
    editor.remove_slide_table_row(0, table.model_object_id, RowDeletion::header(0))?;
    editor.remove_slide_table_row(0, table.model_object_id, RowDeletion::footer(1))?;
    editor.remove_slide_table_column(0, table.model_object_id, ColumnDeletion::header(0))?;
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

fn set_numbers_cells(
    editor: &mut NumbersEditor,
    updates: impl IntoIterator<Item = litchi_numbers::cell::Update>,
) -> Result<(), Box<dyn std::error::Error>> {
    let changes = updates
        .into_iter()
        .map(numbers_cell_change)
        .collect::<Result<Vec<_>, _>>()?;
    let package = litchi_numbers::Package::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_table_cells(
            litchi_numbers::SheetSelector::index(0),
            litchi_numbers::TableSelector::index(0),
        )?
        .extend(changes)?
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    *editor = NumbersEditor::from_bytes(&bytes)?;
    Ok(())
}

fn numbers_cell_change(
    update: litchi_numbers::cell::Update,
) -> Result<litchi_numbers::table::cells::Change, Box<dyn std::error::Error>> {
    let position = litchi_numbers::CellPosition::try_from_usize(update.row, update.column)?;
    let change = match update.value {
        CellValue::Empty => litchi_numbers::table::cells::Change::clear(position),
        CellValue::Text(value) => litchi_numbers::table::cells::Change::set(
            position,
            litchi_numbers::table::cells::Input::text(value)?,
        ),
        CellValue::Number(value) => litchi_numbers::table::cells::Change::set(
            position,
            litchi_numbers::table::cells::Input::number(value.get())?,
        ),
        CellValue::Boolean(value) => litchi_numbers::table::cells::Change::set(
            position,
            litchi_numbers::table::cells::Input::boolean(value),
        ),
        CellValue::Date(value) => litchi_numbers::table::cells::Change::set(
            position,
            litchi_numbers::table::cells::Input::date(value.get())?,
        ),
        CellValue::Duration(value) => litchi_numbers::table::cells::Change::set(
            position,
            litchi_numbers::table::cells::Input::duration(value.get())?,
        ),
        CellValue::Formula(_) | CellValue::Error(_) => {
            return Err(std::io::Error::other("unsupported Numbers cell input").into());
        },
    };
    Ok(change)
}
