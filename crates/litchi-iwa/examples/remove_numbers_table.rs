//! Remove a table from its owning Numbers sheet.

use std::env;

use litchi_iwa::numbers::NumbersEditor;
use litchi_numbers::TableSelector;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: remove_numbers_table <input.numbers> <output.numbers> <table-index>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let table_index: usize = arguments.next().ok_or("missing table index")?.parse()?;

    let mut editor = NumbersEditor::open(input)?;
    editor
        .tables()?
        .get(table_index)
        .ok_or("table index is out of bounds")?;
    editor.remove_table(TableSelector::index(table_index))?;
    editor.save(output)?;
    Ok(())
}
