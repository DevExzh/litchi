//! Create, update, or delete a Numbers cell comment.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_numbers_comment <input.numbers> <output.numbers> <table-id-or-name> <row> <column> <text|--clear>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let table_selector = arguments.next().ok_or("missing table ID or name")?;
    let row = arguments
        .next()
        .ok_or("missing zero-based row")?
        .parse::<usize>()?;
    let column = arguments
        .next()
        .ok_or("missing zero-based column")?
        .parse::<usize>()?;
    let replacement = arguments.collect::<Vec<_>>().join(" ");
    if replacement.is_empty() {
        return Err("missing comment text or --clear".into());
    }

    let mut editor = NumbersEditor::open(&input)?;
    let tables = editor.tables()?;
    let table = table_selector
        .parse::<u64>()
        .ok()
        .and_then(|id| tables.iter().find(|table| table.id() == id))
        .or_else(|| tables.iter().find(|table| table.name == table_selector));
    let table_id = table
        .ok_or("table selector did not match a Numbers table")?
        .id();

    if replacement == "--clear" {
        editor.clear_cell_comment(table_id, row, column)?;
    } else {
        editor.set_cell_comment(table_id, row, column, replacement)?;
    }
    editor.save(&output)?;

    let verified = NumbersEditor::open(&output)?.cell_comment(table_id, row, column)?;
    println!("comment={verified:?}");
    Ok(())
}
