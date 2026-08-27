use std::env;

use litchi_iwa::keynote::KeynoteEditor;
use litchi_keynote::slide::table::dimension::Dimension;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: list_keynote_tables <presentation.key>")?;
    let editor = KeynoteEditor::open(input)?;
    let package = litchi_keynote::Package::from_bytes(&editor.to_bytes()?)?;
    for slide in editor.slides()? {
        for (table_index, info) in editor.slide_tables(slide.index)?.into_iter().enumerate() {
            let table = editor.slide_table(slide.index, info.model_object_id)?;
            let headers = package.slide_table_header_settings(
                litchi_keynote::SlideSelector::index(slide.index),
                litchi_keynote::TableSelector::index(table_index),
            )?;
            let title = package.slide_table_title_settings(
                litchi_keynote::SlideSelector::index(slide.index),
                litchi_keynote::TableSelector::index(table_index),
            )?;
            let row_heights = (0..info.rows)
                .map(|row| {
                    package.slide_table_dimension_size(
                        litchi_keynote::SlideSelector::index(slide.index),
                        litchi_keynote::TableSelector::index(table_index),
                        Dimension::Row(row),
                    )
                })
                .collect::<Result<Vec<_>, _>>()?;
            let column_widths = (0..info.columns)
                .map(|column| {
                    package.slide_table_dimension_size(
                        litchi_keynote::SlideSelector::index(slide.index),
                        litchi_keynote::TableSelector::index(table_index),
                        Dimension::Column(column),
                    )
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
