use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_pages_tables INPUT")?;
    let editor = PagesEditor::open(input)?;
    for table in editor.tables()? {
        let headers = editor.table_header_settings(table.model_object_id)?;
        let row_heights = (0..table.rows)
            .map(|row| editor.table_row_height(table.model_object_id, row))
            .collect::<Result<Vec<_>, _>>()?;
        let column_widths = (0..table.columns)
            .map(|column| editor.table_column_width(table.model_object_id, column))
            .collect::<Result<Vec<_>, _>>()?;
        println!(
            "anchor={} drawable={} model={} name={:?} dimensions={}x{} headers={headers:?} row_heights={row_heights:?} column_widths={column_widths:?}",
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
