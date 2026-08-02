//! Resize an existing Numbers table by table index.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: resize_numbers_table <input.numbers> <output.numbers> <table-index> <rows> <columns>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let table_index: usize = arguments.next().ok_or("missing table index")?.parse()?;
    let rows: usize = arguments.next().ok_or("missing row count")?.parse()?;
    let columns: usize = arguments.next().ok_or("missing column count")?.parse()?;

    let mut editor = NumbersEditor::open(input)?;
    let table = editor
        .tables()?
        .get(table_index)
        .cloned()
        .ok_or("table index is out of bounds")?;
    editor.resize_table(table.object_id, rows, columns)?;
    editor.save(output)?;
    Ok(())
}
