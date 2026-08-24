use std::env;

use litchi_iwa::pages::{
    PagesCellValue, PagesDocumentBuilder, PagesEditor, PagesTableCellUpdate,
    PagesTableDimensionSize, PagesTableFormulaCachedValue, PagesTableFormulaCellReference,
    PagesTableFormulaExpression,
};
use litchi_numbers::table::topology::{ColumnInsertion, RowInsertion};
use litchi_pages::table::headers::{Count as HeaderCount, Settings as HeaderSettings};
use litchi_pages::table::title::Settings as PagesTableTitleSettings;
use litchi_pages::{BodyTableSelector, Package};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = env::args()
        .nth(1)
        .unwrap_or_else(|| "scratch-table.pages".to_owned());
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Quarterly revenue\n")
        .body_table("Revenue", 4, 3)
        .build()?;
    let table = editor.tables()?.remove(0);
    editor.set_table_cells(
        table.model_object_id,
        [
            PagesTableCellUpdate::new(0, 0, PagesCellValue::Text("Quarter".to_owned())),
            PagesTableCellUpdate::new(0, 1, PagesCellValue::Text("Revenue".to_owned())),
            PagesTableCellUpdate::new(0, 2, PagesCellValue::Text("Growth".to_owned())),
            PagesTableCellUpdate::new(1, 0, PagesCellValue::Text("Q1".to_owned())),
            PagesTableCellUpdate::new(1, 1, PagesCellValue::number(125_000.0)?),
            PagesTableCellUpdate::new(1, 2, PagesCellValue::number(0.18)?),
        ],
    )?;
    editor.set_table_formula(
        table.model_object_id,
        3,
        1,
        PagesTableFormulaExpression::function(
            "SUM",
            [PagesTableFormulaExpression::range(
                PagesTableFormulaCellReference::relative(1, 1),
                PagesTableFormulaCellReference::relative(2, 1),
            )],
        ),
        PagesTableFormulaCachedValue::Number(125_000.0.try_into()?),
    )?;
    let focused = Package::from_bytes(&editor.to_bytes()?)?;
    let header_commit = focused
        .edit_body_table_header_settings(BodyTableSelector::name("Revenue"))?
        .set(HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            header_columns: Some(HeaderCount::ONE),
            footer_rows: Some(HeaderCount::ONE),
            ..Default::default()
        })
        .commit()?;
    let mut header_bytes = Vec::new();
    header_commit.package().write_to(&mut header_bytes)?;
    editor = PagesEditor::from_bytes(&header_bytes)?;
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
    editor.insert_table_row(table.model_object_id, RowInsertion::body(2))?;
    editor.insert_table_column(table.model_object_id, ColumnInsertion::body(1))?;
    editor.set_table_cells(
        table.model_object_id,
        [
            PagesTableCellUpdate::new(3, 0, PagesCellValue::Text("Q2".to_owned())),
            PagesTableCellUpdate::new(3, 1, PagesCellValue::number(142_000.0)?),
            PagesTableCellUpdate::new(3, 2, PagesCellValue::Text("provisional".to_owned())),
            PagesTableCellUpdate::new(3, 3, PagesCellValue::number(0.14)?),
            PagesTableCellUpdate::new(0, 2, PagesCellValue::Text("Status".to_owned())),
        ],
    )?;
    let second_anchor = editor.body_text()?.encode_utf16().count();
    let notes = editor.add_table(second_anchor, "Notes", 2, 2)?;
    editor.set_table_cell(
        notes.model_object_id,
        0,
        0,
        PagesCellValue::Text("Generated independently".to_owned()),
    )?;
    let focused = Package::from_bytes(&editor.to_bytes()?)?;
    let commit = focused
        .edit_body_table_title(BodyTableSelector::name("Revenue"))?
        .set(PagesTableTitleSettings::new(Some(true), Some(true)))
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    std::fs::write(output, bytes)?;
    Ok(())
}
