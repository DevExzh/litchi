//! Create Pages, Numbers, and Keynote files with native table-cell fills.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeFill};
use litchi_numbers::cell::Value as CellValue;

const ROW: usize = 1;
const COLUMN: usize = 1;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_table_fills <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;
    create_numbers(&output.join("table-fills.numbers"))?;
    create_pages(&output.join("table-fills.pages"))?;
    create_keynote(&output.join("table-fills.key"))?;
    Ok(())
}

fn fill() -> Result<ShapeFill, Box<dyn std::error::Error>> {
    Ok(ShapeFill::Solid(RgbaColor::new(
        0.96,
        0.72,
        0.12,
        1.0,
        RgbColorSpace::Srgb,
    )?))
}

fn create_numbers(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Fills")
        .table_dimensions(3, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).id();
    set_numbers_cell(
        &mut editor,
        ROW,
        COLUMN,
        CellValue::Text("Numbers".to_owned()),
    )?;
    editor.set_table_cell_fill(table_id, ROW, COLUMN, &fill()?)?;
    editor.save(output)?;
    Ok(())
}

fn set_numbers_cell(
    editor: &mut NumbersEditor,
    row: usize,
    column: usize,
    value: CellValue,
) -> Result<(), Box<dyn std::error::Error>> {
    set_numbers_cells(
        editor,
        [litchi_numbers::cell::Update::new(row, column, value)],
    )
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

fn create_pages(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Created from scratch with a native table-cell fill.\n")
        .body_table("Fills", 3, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.set_table_cell(table_id, ROW, COLUMN, CellValue::Text("Pages".to_owned()))?;
    editor.set_table_cell_fill(table_id, ROW, COLUMN, &fill()?)?;
    editor.save(output)?;
    Ok(())
}

fn create_keynote(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Native table-cell fills")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Fills",
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
        CellValue::Text("Keynote".to_owned()),
    )?;
    editor.set_slide_table_cell_fill(0, table.model_object_id, ROW, COLUMN, &fill()?)?;
    editor.save(output)?;
    Ok(())
}
