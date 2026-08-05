//! Create Pages, Numbers, and Keynote documents with hidden table axes.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::{PagesDocumentBuilder, PagesEditor};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::table::axis::{AxisIndex, HiddenAxes};
use litchi_numbers::cell::{Update as TableCellUpdate, Value as CellValue};

const TABLE_ROWS: usize = 4;
const TABLE_COLUMNS: usize = 3;
const HIDDEN_ROW: usize = 2;
const HIDDEN_COLUMN: usize = 1;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_hidden_tables <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;
    create_numbers(&output.join("hidden-table.numbers"))?;
    create_pages(&output.join("hidden-table.pages"))?;
    create_keynote(&output.join("hidden-table.key"))?;
    Ok(())
}

fn create_numbers(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let hidden = HiddenAxes::new([AxisIndex::row(HIDDEN_ROW), AxisIndex::column(HIDDEN_COLUMN)])?;
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Hidden Axes")
        .table_dimensions(TABLE_ROWS, TABLE_COLUMNS)
        .build()?;
    let table_id = editor.tables()?.remove(0).object_id;
    editor.set_cells(table_id, table_cells())?;
    editor.set_table_hidden_axes(table_id, &hidden)?;
    editor.save(output)?;
    if NumbersEditor::open(output)?.table_hidden_axes(table_id)? != hidden {
        return Err("Numbers hidden axes failed reopen validation".into());
    }
    Ok(())
}

fn create_pages(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let hidden = HiddenAxes::new([AxisIndex::row(HIDDEN_ROW), AxisIndex::column(HIDDEN_COLUMN)])?;
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Hidden table axes created by litchi-iwa\n")
        .body_table("Hidden Axes", TABLE_ROWS, TABLE_COLUMNS)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.set_table_cells(table_id, table_cells())?;
    editor.set_table_hidden_axes(table_id, &hidden)?;
    editor.save(output)?;
    if PagesEditor::open(output)?.table_hidden_axes(table_id)? != hidden {
        return Err("Pages hidden axes failed reopen validation".into());
    }
    Ok(())
}

fn create_keynote(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let hidden = HiddenAxes::new([AxisIndex::row(HIDDEN_ROW), AxisIndex::column(HIDDEN_COLUMN)])?;
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Hidden table axes")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Hidden Axes",
        TABLE_ROWS,
        TABLE_COLUMNS,
        DrawablePoint { x: 320.0, y: 360.0 },
        DrawableSize {
            width: 1_280.0,
            height: 480.0,
        },
    )?;
    editor.set_slide_table_cells(0, table.model_object_id, table_cells())?;
    editor.set_slide_table_hidden_axes(0, table.model_object_id, &hidden)?;
    editor.save(output)?;
    if KeynoteEditor::open(output)?.slide_table_hidden_axes(0, table.model_object_id)? != hidden {
        return Err("Keynote hidden axes failed reopen validation".into());
    }
    Ok(())
}

fn table_cells() -> [TableCellUpdate; 12] {
    [
        TableCellUpdate::new(0, 0, CellValue::Text("Name".to_owned())),
        TableCellUpdate::new(0, 1, CellValue::Text("Private".to_owned())),
        TableCellUpdate::new(0, 2, CellValue::Text("Status".to_owned())),
        TableCellUpdate::new(1, 0, CellValue::Text("Alpha".to_owned())),
        TableCellUpdate::new(1, 1, CellValue::Text("x".to_owned())),
        TableCellUpdate::new(1, 2, CellValue::Text("Ready".to_owned())),
        TableCellUpdate::new(2, 0, CellValue::Text("Hidden row".to_owned())),
        TableCellUpdate::new(2, 1, CellValue::Text("y".to_owned())),
        TableCellUpdate::new(2, 2, CellValue::Text("Hidden".to_owned())),
        TableCellUpdate::new(3, 0, CellValue::Text("Omega".to_owned())),
        TableCellUpdate::new(3, 1, CellValue::Text("z".to_owned())),
        TableCellUpdate::new(3, 2, CellValue::Text("Done".to_owned())),
    ]
}
