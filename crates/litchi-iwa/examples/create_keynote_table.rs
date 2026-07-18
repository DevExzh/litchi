use std::env;

use litchi_iwa::keynote::{
    KeynoteDocumentBuilder, KeynoteTableCellUpdate, KeynoteTableCellValue,
    KeynoteTableDimensionSize, KeynoteTableFormulaCachedValue, KeynoteTableFormulaCellReference,
    KeynoteTableFormulaExpression, KeynoteTableHeaderCount, KeynoteTableHeaderSettings,
    KeynoteTableTitleSettings,
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
    let mut updates = Vec::new();
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
            updates.push(KeynoteTableCellUpdate::new(row, column, value));
        }
    }
    editor.set_slide_table_cells(0, table.model_object_id, updates)?;
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
    editor.set_slide_table_title_settings(
        0,
        table.model_object_id,
        KeynoteTableTitleSettings {
            visible: Some(true),
            outlined: Some(true),
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
    editor.set_slide_table_formula(
        0,
        table.model_object_id,
        3,
        1,
        KeynoteTableFormulaExpression::function(
            "SUM",
            [KeynoteTableFormulaExpression::range(
                KeynoteTableFormulaCellReference::relative(1, 1),
                KeynoteTableFormulaCellReference::relative(2, 1),
            )],
        ),
        KeynoteTableFormulaCachedValue::Number(218.0),
    )?;
    editor.insert_slide_table_row(0, table.model_object_id, 3)?;
    editor.insert_slide_table_column(0, table.model_object_id, 2)?;
    editor.set_slide_table_cells(
        0,
        table.model_object_id,
        [
            KeynoteTableCellUpdate::new(3, 0, KeynoteTableCellValue::Text("Central".to_owned())),
            KeynoteTableCellUpdate::new(3, 1, KeynoteTableCellValue::Number(105.0)),
            KeynoteTableCellUpdate::new(3, 2, KeynoteTableCellValue::Text("review".to_owned())),
            KeynoteTableCellUpdate::new(3, 3, KeynoteTableCellValue::Number(139.0)),
            KeynoteTableCellUpdate::new(0, 2, KeynoteTableCellValue::Text("Status".to_owned())),
        ],
    )?;
    editor.save(&output)?;
    println!("created {output}");
    Ok(())
}
