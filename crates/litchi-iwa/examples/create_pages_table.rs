use std::env;

use litchi_iwa::pages::{
    PagesCellValue, PagesDocumentBuilder, PagesTableDimensionSize, PagesTableFormulaCachedValue,
    PagesTableFormulaExpression, PagesTableHeaderCount, PagesTableHeaderSettings,
    PagesTableTitleSettings,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = env::args()
        .nth(1)
        .unwrap_or_else(|| "scratch-table.pages".to_owned());
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Quarterly revenue\n")
        .body_table("Revenue", 4, 3)
        .build()?;
    let table = editor.tables()?.remove(0);
    for (column, heading) in ["Quarter", "Revenue", "Growth"].into_iter().enumerate() {
        editor.set_table_cell(
            table.model_object_id,
            0,
            column,
            PagesCellValue::Text(heading.to_owned()),
        )?;
    }
    editor.set_table_cell(
        table.model_object_id,
        1,
        0,
        PagesCellValue::Text("Q1".to_owned()),
    )?;
    editor.set_table_cell(
        table.model_object_id,
        1,
        1,
        PagesCellValue::Number(125_000.0),
    )?;
    editor.set_table_cell(table.model_object_id, 1, 2, PagesCellValue::Number(0.18))?;
    editor.set_table_formula(
        table.model_object_id,
        2,
        1,
        PagesTableFormulaExpression::function(
            "SUM",
            [
                PagesTableFormulaExpression::Number(1.0),
                PagesTableFormulaExpression::Number(2.0),
            ],
        ),
        PagesTableFormulaCachedValue::Number(3.0),
    )?;
    editor.set_table_header_settings(
        table.model_object_id,
        PagesTableHeaderSettings {
            header_rows: Some(PagesTableHeaderCount::ONE),
            header_columns: Some(PagesTableHeaderCount::ONE),
            footer_rows: Some(PagesTableHeaderCount::ONE),
            ..Default::default()
        },
    )?;
    editor.set_table_title_settings(
        table.model_object_id,
        PagesTableTitleSettings {
            visible: Some(true),
            outlined: Some(true),
        },
    )?;
    for (column, width) in [120.0, 160.0, 100.0].into_iter().enumerate() {
        editor.set_table_column_width(
            table.model_object_id,
            column,
            PagesTableDimensionSize::points(width)?,
        )?;
    }
    for (row, height) in [28.0, 34.0, 40.0, 46.0].into_iter().enumerate() {
        editor.set_table_row_height(
            table.model_object_id,
            row,
            PagesTableDimensionSize::points(height)?,
        )?;
    }
    editor.insert_table_row(table.model_object_id, 2)?;
    editor.insert_table_column(table.model_object_id, 2)?;
    for (column, value) in [
        PagesCellValue::Text("Q2".to_owned()),
        PagesCellValue::Number(142_000.0),
        PagesCellValue::Text("provisional".to_owned()),
        PagesCellValue::Number(0.14),
    ]
    .into_iter()
    .enumerate()
    {
        editor.set_table_cell(table.model_object_id, 2, column, value)?;
    }
    editor.set_table_cell(
        table.model_object_id,
        0,
        2,
        PagesCellValue::Text("Status".to_owned()),
    )?;
    let second_anchor = editor.body_text()?.encode_utf16().count();
    let notes = editor.add_table(second_anchor, "Notes", 2, 2)?;
    editor.set_table_cell(
        notes.model_object_id,
        0,
        0,
        PagesCellValue::Text("Generated independently".to_owned()),
    )?;
    editor.save(output)?;
    Ok(())
}
