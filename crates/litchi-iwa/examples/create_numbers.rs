//! Create a Numbers spreadsheet without an input document or template.

use litchi_iwa::numbers::{
    FormulaCachedValue, FormulaCellReference, FormulaExpression, NumbersDocumentBuilder,
    NumbersTableHeaderCount, NumbersTableHeaderSettings, TableRowInsertion,
};
use litchi_iwa::text::{Font, TextStyle};
use litchi_numbers::cell::{Update as TableCellUpdate, Value as CellValue};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_numbers <output.numbers>")?;
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Forecast")
        .table_dimensions(4, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).object_id;
    let mut updates = Vec::new();
    for (column, heading) in ["Region", "Q1", "Q2"].into_iter().enumerate() {
        updates.push(TableCellUpdate::new(
            0,
            column,
            CellValue::Text(heading.to_owned()),
        ));
    }
    for (row, region, q1, q2) in [(1, "North", 120.0, 145.0), (2, "South", 98.0, 132.0)] {
        updates.extend([
            TableCellUpdate::new(row, 0, CellValue::Text(region.to_owned())),
            TableCellUpdate::new(row, 1, CellValue::Number(q1)),
            TableCellUpdate::new(row, 2, CellValue::Number(q2)),
        ]);
    }
    editor.set_cells(table_id, updates)?;
    editor.set_table_cell_text_style(table_id, 1, 0, TextStyle::default().with_bold(true))?;
    editor.set_table_cell_text_font(table_id, 1, 0, Font::named("CourierNewPSMT")?)?;
    editor.set_table_header_settings(
        table_id,
        NumbersTableHeaderSettings {
            header_rows: Some(NumbersTableHeaderCount::ONE),
            header_columns: Some(NumbersTableHeaderCount::ONE),
            footer_rows: Some(NumbersTableHeaderCount::ONE),
            ..Default::default()
        },
    )?;
    editor.set_formula_with_cached_value(
        table_id,
        3,
        1,
        FormulaExpression::function(
            "SUM",
            [FormulaExpression::range(
                FormulaCellReference::relative(1, 1),
                FormulaCellReference::relative(2, 1),
            )],
        ),
        FormulaCachedValue::Number(218.0),
    )?;
    editor.insert_table_row(table_id, TableRowInsertion::body(2))?;
    editor.set_cells(
        table_id,
        [
            TableCellUpdate::new(3, 0, CellValue::Text("Central".to_owned())),
            TableCellUpdate::new(3, 1, CellValue::Number(105.0)),
            TableCellUpdate::new(3, 2, CellValue::Number(139.0)),
        ],
    )?;
    editor.save(output)?;
    Ok(())
}
