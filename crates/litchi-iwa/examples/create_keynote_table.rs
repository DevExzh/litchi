use std::env;

use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteTableCellValue};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = env::args()
        .nth(1)
        .unwrap_or_else(|| "keynote-table.key".to_owned());
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Regional outlook")
        .subtitle("Created entirely by litchi-iwa")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Forecast",
        4,
        3,
        DrawablePoint { x: 350.0, y: 390.0 },
        DrawableSize {
            width: 1_220.0,
            height: 430.0,
        },
    )?;
    for (row, values) in [
        ["Region", "Q1", "Q2"],
        ["North", "120", "145"],
        ["South", "98", "132"],
        ["West", "110", "127"],
    ]
    .into_iter()
    .enumerate()
    {
        for (column, value) in values.into_iter().enumerate() {
            let value = if row == 0 || column == 0 {
                KeynoteTableCellValue::Text(value.to_owned())
            } else {
                KeynoteTableCellValue::Number(value.parse()?)
            };
            editor.set_slide_table_cell(0, table.model_object_id, row, column, value)?;
        }
    }
    editor.save(&output)?;
    println!("created {output}");
    Ok(())
}
