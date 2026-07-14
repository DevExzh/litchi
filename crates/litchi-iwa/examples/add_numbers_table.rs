//! Add an independent empty table to an existing Numbers sheet.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: add_numbers_table <input.numbers> <output.numbers> <sheet-index> <name> <rows> <columns>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet_index: usize = arguments.next().ok_or("missing sheet index")?.parse()?;
    let name = arguments.next().ok_or("missing table name")?;
    let rows = arguments.next().ok_or("missing row count")?.parse()?;
    let columns = arguments.next().ok_or("missing column count")?.parse()?;

    let mut editor = NumbersEditor::open(input)?;
    let sheet_id = editor
        .sheets()?
        .get(sheet_index)
        .ok_or("sheet index is out of range")?
        .object_id;
    let table = editor.add_empty_table(sheet_id, &name, rows, columns)?;
    editor.save(output)?;
    println!(
        "created table {} ({:?}) with {} rows and {} columns",
        table.object_id, table.name, table.rows, table.columns
    );
    Ok(())
}
