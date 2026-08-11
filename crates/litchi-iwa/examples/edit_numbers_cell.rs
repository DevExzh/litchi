//! Set one primitive cell value in a Numbers spreadsheet.

use std::env;

use litchi_iwa::numbers::NumbersEditor;
use litchi_numbers::cell::Value as CellValue;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_numbers_cell <input> <output> <table-index> <row> <column> <type> [value]",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let table_index: usize = arguments.next().ok_or("missing table index")?.parse()?;
    let row: usize = arguments.next().ok_or("missing zero-based row")?.parse()?;
    let column: usize = arguments
        .next()
        .ok_or("missing zero-based column")?
        .parse()?;
    let value_type = arguments.next().ok_or("missing value type")?;
    let raw_value = arguments.next();

    let value = match value_type.as_str() {
        "empty" => CellValue::Empty,
        "text" => CellValue::Text(raw_value.ok_or("missing text value")?),
        "number" => CellValue::number(raw_value.ok_or("missing number value")?.parse()?)?,
        "boolean" => CellValue::Boolean(raw_value.ok_or("missing boolean value")?.parse()?),
        "date" => CellValue::date(raw_value.ok_or("missing date value")?.parse()?)?,
        "duration" => CellValue::duration(raw_value.ok_or("missing duration value")?.parse()?)?,
        _ => return Err(format!("unsupported value type {value_type:?}").into()),
    };

    let mut editor = NumbersEditor::open(input)?;
    let tables = editor.tables()?;
    let table = tables
        .get(table_index)
        .ok_or("table index is out of bounds")?;
    editor.set_cell(table.id(), row, column, value)?;
    editor.save(output)?;
    Ok(())
}
