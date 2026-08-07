use std::env;
use std::path::PathBuf;

use litchi_iwa::numbers::NumbersEditor;
use litchi_numbers::TableSelector;
use litchi_numbers::table::topology::RowDeletion;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input =
        PathBuf::from(arguments.next().ok_or(
            "usage: remove_numbers_row <input.numbers> <output.numbers> <table-index> <header|body|footer> <index>",
        )?);
    let output = PathBuf::from(arguments.next().ok_or("missing output path")?);
    let table_index = arguments
        .next()
        .ok_or("missing table index")?
        .parse::<usize>()?;
    let section = arguments.next().ok_or("missing row section")?;
    let index = arguments
        .next()
        .ok_or("missing section-relative row index")?
        .parse::<usize>()?;
    if arguments.next().is_some() {
        return Err("unexpected extra argument".into());
    }

    let deletion = match section.as_str() {
        "header" => RowDeletion::header(index),
        "body" => RowDeletion::body(index),
        "footer" => RowDeletion::footer(index),
        _ => return Err(format!("unsupported row section {section:?}").into()),
    };

    let mut editor = NumbersEditor::open(input)?;
    editor.remove_table_row(TableSelector::index(table_index), deletion)?;
    editor.save(output)?;
    Ok(())
}
