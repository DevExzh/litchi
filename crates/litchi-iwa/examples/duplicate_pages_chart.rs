//! Duplicate a native Pages body chart at a UTF-16 body position.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: duplicate_pages_chart <input.pages> <output.pages> <chart-index> <utf16-anchor>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let chart_index: usize = arguments.next().ok_or("missing chart index")?.parse()?;
    let anchor: usize = arguments.next().ok_or("missing UTF-16 anchor")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = PagesEditor::open(input)?;
    let source = editor
        .body_charts()?
        .into_iter()
        .nth(chart_index)
        .ok_or("chart index is out of bounds")?;
    let duplicate = editor.duplicate_body_chart(source.drawable_object_id, anchor)?;
    editor.save(output)?;
    println!(
        "anchor={} drawable={} source={} rows={} columns={}",
        duplicate.anchor_character_index,
        duplicate.drawable_object_id,
        source.drawable_object_id,
        duplicate.data.row_names().len(),
        duplicate.data.column_names().len(),
    );
    Ok(())
}
