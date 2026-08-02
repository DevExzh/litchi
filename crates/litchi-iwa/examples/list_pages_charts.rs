//! List native charts anchored to a Pages document body.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: list_pages_charts <input.pages>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let editor = PagesEditor::open(input)?;
    for chart in editor.body_charts()? {
        println!(
            "anchor={} object={} kind={:?} direction={:?} rows={} columns={}",
            chart.anchor_character_index,
            chart.drawable_object_id,
            chart.kind,
            chart.direction,
            chart.data.row_names().len(),
            chart.data.column_names().len(),
        );
    }
    Ok(())
}
