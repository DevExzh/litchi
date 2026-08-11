//! Move a populated table between existing Numbers sheets.

use std::env;

use litchi_iwa::numbers::NumbersEditor;
use litchi_numbers::{SheetSelector, TableSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: move_numbers_table <input.numbers> <output.numbers> <table-index> <target-sheet-index>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let table_index: usize = arguments.next().ok_or("missing table index")?.parse()?;
    let target_sheet_index: usize = arguments
        .next()
        .ok_or("missing target sheet index")?
        .parse()?;

    let mut editor = NumbersEditor::open(input)?;
    let table = editor
        .tables()?
        .get(table_index)
        .cloned()
        .ok_or_else(|| format!("table index {table_index} is out of range"))?;
    let target = editor
        .sheets()?
        .get(target_sheet_index)
        .cloned()
        .ok_or_else(|| format!("sheet index {target_sheet_index} is out of range"))?;
    editor.move_table(
        TableSelector::index(table_index),
        SheetSelector::index(target_sheet_index),
    )?;
    editor.save(output)?;

    println!("moved table {:?} to sheet {:?}", table.name, target.name);
    Ok(())
}
