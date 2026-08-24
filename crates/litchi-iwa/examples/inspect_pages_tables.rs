use std::env;

use litchi_iwa::pages::PagesEditor;
use litchi_pages::table::dimension::Dimension;
use litchi_pages::{BodyTableSelector, Package};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_pages_tables INPUT")?;
    let editor = PagesEditor::open(&input)?;
    let focused = Package::open(&input)?;
    for (position, table) in editor.tables()?.into_iter().enumerate() {
        let materialized = editor.table(table.model_object_id)?;
        let cells = materialized.iter_cells().collect::<Vec<_>>();
        let headers = focused.body_table_header_settings(BodyTableSelector::index(position))?;
        let title =
            focused.body_table_title_settings(BodyTableSelector::position(position.into()))?;
        let row_heights = (0..table.rows)
            .map(|row| {
                focused.body_table_dimension_size(
                    BodyTableSelector::index(position),
                    Dimension::Row(row),
                )
            })
            .collect::<Result<Vec<_>, _>>()?;
        let column_widths = (0..table.columns)
            .map(|column| {
                focused.body_table_dimension_size(
                    BodyTableSelector::index(position),
                    Dimension::Column(column),
                )
            })
            .collect::<Result<Vec<_>, _>>()?;
        println!(
            "anchor={} drawable={} model={} name={:?} dimensions={}x{} title={title:?} headers={headers:?} row_heights={row_heights:?} column_widths={column_widths:?} cells={cells:?}",
            table.anchor_character_index,
            table.drawable_object_id,
            table.model_object_id,
            table.name,
            table.rows,
            table.columns,
        );
    }
    Ok(())
}
