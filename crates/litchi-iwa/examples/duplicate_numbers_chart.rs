//! Duplicate a native Numbers sheet chart.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: duplicate_numbers_chart <input.numbers> <output.numbers> <sheet-index> <chart-index>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet_index: usize = arguments.next().ok_or("missing sheet index")?.parse()?;
    let chart_index: usize = arguments.next().ok_or("missing chart index")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersEditor::open(input)?;
    let sheet = editor
        .sheets()?
        .into_iter()
        .nth(sheet_index)
        .ok_or("sheet index is out of bounds")?;
    let source = editor
        .sheet_charts(sheet.id())?
        .into_iter()
        .nth(chart_index)
        .ok_or("chart index is out of bounds")?;
    let duplicate = editor.duplicate_sheet_chart(sheet.id(), source.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "sheet={} drawable={} source={} rows={} columns={}",
        sheet.id(),
        duplicate.drawable_object_id,
        source.drawable_object_id,
        duplicate.data.row_names().len(),
        duplicate.data.column_names().len(),
    );
    Ok(())
}
