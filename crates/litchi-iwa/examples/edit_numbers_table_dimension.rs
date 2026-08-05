use std::env;
use std::path::PathBuf;

use litchi_iwa::numbers::{Dimension, NumbersEditor, Size};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = PathBuf::from(arguments.next().ok_or(
        "usage: edit_numbers_table_dimension <input.numbers> <output.numbers> <table-id> <row|column> <index> <default|points>",
    )?);
    let output = PathBuf::from(arguments.next().ok_or("missing output path")?);
    let table_id = arguments.next().ok_or("missing table ID")?.parse::<u64>()?;
    let axis = arguments.next().ok_or("missing dimension axis")?;
    let index = arguments
        .next()
        .ok_or("missing dimension index")?
        .parse::<usize>()?;
    let value = arguments.next().ok_or("missing dimension size")?;
    if arguments.next().is_some() {
        return Err("unexpected extra argument".into());
    }
    let dimension = match axis.as_str() {
        "row" => Dimension::Row(index),
        "column" => Dimension::Column(index),
        _ => return Err("dimension axis must be row or column".into()),
    };
    let size = if value == "default" {
        Size::Default
    } else {
        Size::points(value.parse::<f32>()?)?
    };

    let mut editor = NumbersEditor::open(input)?;
    editor.set_table_dimension_size(table_id, dimension, size)?;
    editor.save(output)?;
    Ok(())
}
