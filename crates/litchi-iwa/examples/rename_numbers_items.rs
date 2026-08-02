//! Rename a Numbers sheet and table by their stable object IDs.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: rename_numbers_items <input.numbers> <output.numbers> <sheet-name> <table-name>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet_name = arguments.next().ok_or("missing sheet name")?;
    let table_name = arguments.next().ok_or("missing table name")?;

    let mut editor = NumbersEditor::open(input)?;
    let sheet = editor.sheets()?.into_iter().next().ok_or("no sheets")?;
    let table = editor.tables()?.into_iter().next().ok_or("no tables")?;
    editor.rename_sheet(sheet.object_id, &sheet_name)?;
    editor.rename_table(table.object_id, &table_name)?;
    editor.save(output)?;
    Ok(())
}
