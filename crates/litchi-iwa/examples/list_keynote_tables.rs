use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: list_keynote_tables <presentation.key>")?;
    let editor = KeynoteEditor::open(input)?;
    for slide in editor.slides()? {
        for info in editor.slide_tables(slide.index)? {
            let table = editor.slide_table(slide.index, info.model_object_id)?;
            let headers = editor.slide_table_header_settings(slide.index, info.model_object_id)?;
            let title = editor.slide_table_title_settings(slide.index, info.model_object_id)?;
            let row_heights = (0..info.rows)
                .map(|row| editor.slide_table_row_height(slide.index, info.model_object_id, row))
                .collect::<Result<Vec<_>, _>>()?;
            let column_widths = (0..info.columns)
                .map(|column| {
                    editor.slide_table_column_width(slide.index, info.model_object_id, column)
                })
                .collect::<Result<Vec<_>, _>>()?;
            println!(
                "slide={} drawable={} model={} name={:?} rows={} columns={} title={:?} headers={:?} row_heights={:?} column_widths={:?} cells={:?}",
                slide.index + 1,
                info.drawable_object_id,
                info.model_object_id,
                info.name,
                info.rows,
                info.columns,
                title,
                headers,
                row_heights,
                column_widths,
                table.iter_cells().collect::<Vec<_>>()
            );
        }
    }
    Ok(())
}
