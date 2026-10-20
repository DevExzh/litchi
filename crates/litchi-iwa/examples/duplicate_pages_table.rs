use std::env;

use litchi_iwa::pages::{self};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = env::args()
        .nth(1)
        .unwrap_or_else(|| "pages-duplicated-table.pages".to_owned());
    let mut editor = pages::Builder::new()
        .body_text("Budget scenarios\n")
        .body_table("Budget", 3, 2)
        .build()?;
    let source = editor.tables()?.remove(0);
    editor.set_table_cells(
        source.model_object_id,
        [
            pages::TableCellUpdate::new(0, 0, pages::CellValue::Text("Category".to_owned())),
            pages::TableCellUpdate::new(0, 1, pages::CellValue::Text("Cost".to_owned())),
            pages::TableCellUpdate::new(1, 0, pages::CellValue::Text("Travel".to_owned())),
            pages::TableCellUpdate::new(1, 1, pages::CellValue::number(125.0)?),
        ],
    )?;
    editor.set_table_formula(
        source.model_object_id,
        2,
        1,
        pages::TableFormulaExpression::function(
            "SUM",
            [
                pages::TableFormulaExpression::Number(100.0),
                pages::TableFormulaExpression::Number(25.0),
            ],
        ),
        pages::TableFormulaCachedValue::number(125.0)?,
    )?;

    let anchor = editor.body_text()?.encode_utf16().count();
    let copy = editor.duplicate_table(source.model_object_id, anchor)?;
    editor.set_table_cell(
        copy.model_object_id,
        1,
        0,
        pages::CellValue::Text("Lodging".to_owned()),
    )?;
    editor.save(&output)?;
    println!("created {output} with {} and {}", source.name, copy.name);
    Ok(())
}
