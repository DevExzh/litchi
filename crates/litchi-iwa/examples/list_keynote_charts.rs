//! List native charts owned by every slide in a Keynote presentation.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: list_keynote_charts <input.key>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let editor = KeynoteEditor::open(input)?;
    for slide in editor.slides()? {
        for chart in editor.slide_charts(slide.index)? {
            let legend_font_size =
                editor.slide_chart_legend_font_size(slide.index, chart.drawable_object_id)?;
            println!(
                "slide={} object={} kind={:?} direction={:?} rows={} columns={} legend_font_size={legend_font_size:?}",
                chart.slide_index,
                chart.drawable_object_id,
                chart.kind,
                chart.direction,
                chart.data.row_names().len(),
                chart.data.column_names().len(),
            );
        }
    }
    Ok(())
}
