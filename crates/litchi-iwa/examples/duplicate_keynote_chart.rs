//! Duplicate a native Keynote slide chart.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;
use litchi_keynote::ChartSelector;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: duplicate_keynote_chart <input.key> <output.key> <slide-index> <chart-index>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index: usize = arguments.next().ok_or("missing slide index")?.parse()?;
    let chart_index: usize = arguments.next().ok_or("missing chart index")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = KeynoteEditor::open(input)?;
    let source = editor
        .slide_charts(slide_index)?
        .into_iter()
        .nth(chart_index)
        .ok_or("chart index is out of bounds")?;
    let duplicate = editor.duplicate_slide_chart(slide_index, ChartSelector::index(chart_index))?;
    editor.save(output)?;
    println!(
        "slide={} drawable={} source={} rows={} columns={}",
        slide_index,
        duplicate.drawable_object_id,
        source.drawable_object_id,
        duplicate.data.row_names().len(),
        duplicate.data.column_names().len(),
    );
    Ok(())
}
