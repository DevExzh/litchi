use std::env;

use litchi_iwa::keynote::{
    KeynoteDocumentBuilder, KeynoteTableCellValue, KeynoteTableDimensionSize,
    KeynoteTableHeaderCount, KeynoteTableHeaderSettings,
};
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
    editor.set_slide_table_header_settings(
        0,
        table.model_object_id,
        KeynoteTableHeaderSettings {
            header_rows: Some(KeynoteTableHeaderCount::ONE),
            header_columns: Some(KeynoteTableHeaderCount::ONE),
            footer_rows: Some(KeynoteTableHeaderCount::ONE),
            ..Default::default()
        },
    )?;
    for (column, width) in [440.0, 390.0, 390.0].into_iter().enumerate() {
        editor.set_slide_table_column_width(
            0,
            table.model_object_id,
            column,
            KeynoteTableDimensionSize::points(width)?,
        )?;
    }
    for (row, height) in [90.0, 100.0, 110.0, 130.0].into_iter().enumerate() {
        editor.set_slide_table_row_height(
            0,
            table.model_object_id,
            row,
            KeynoteTableDimensionSize::points(height)?,
        )?;
    }
    editor.save(&output)?;
    println!("created {output}");
    Ok(())
}
