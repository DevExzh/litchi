use std::env;
use std::path::PathBuf;

use litchi_iwa::numbers::NumbersEditor;
use litchi_numbers::TableSelector;
use litchi_numbers::table::topology::ColumnInsertion;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = PathBuf::from(arguments.next().ok_or(
        "usage: insert_numbers_column <input.numbers> <output.numbers> <table-index> <body-column>",
    )?);
    let output = PathBuf::from(arguments.next().ok_or("missing output path")?);
    let table_index = arguments
        .next()
        .ok_or("missing table index")?
        .parse::<usize>()?;
    let column = arguments
        .next()
        .ok_or("missing body column index")?
        .parse::<usize>()?;
    if arguments.next().is_some() {
        return Err("unexpected extra argument".into());
    }

    let mut editor = NumbersEditor::open(input)?;
    editor.insert_table_column(
        TableSelector::index(table_index),
        ColumnInsertion::body(column),
    )?;
    editor.save(output)?;
    Ok(())
}
