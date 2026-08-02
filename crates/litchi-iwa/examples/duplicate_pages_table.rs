use std::env;

use litchi_iwa::pages::{
    PagesCellValue, PagesDocumentBuilder, PagesTableCellUpdate, PagesTableFormulaCachedValue,
    PagesTableFormulaExpression,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = env::args()
        .nth(1)
        .unwrap_or_else(|| "pages-duplicated-table.pages".to_owned());
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Budget scenarios\n")
        .body_table("Budget", 3, 2)
        .build()?;
    let source = editor.tables()?.remove(0);
    editor.set_table_cells(
        source.model_object_id,
        [
            PagesTableCellUpdate::new(0, 0, PagesCellValue::Text("Category".to_owned())),
            PagesTableCellUpdate::new(0, 1, PagesCellValue::Text("Cost".to_owned())),
            PagesTableCellUpdate::new(1, 0, PagesCellValue::Text("Travel".to_owned())),
            PagesTableCellUpdate::new(1, 1, PagesCellValue::Number(125.0)),
        ],
    )?;
    editor.set_table_formula(
        source.model_object_id,
        2,
        1,
        PagesTableFormulaExpression::function(
            "SUM",
            [
                PagesTableFormulaExpression::Number(100.0),
                PagesTableFormulaExpression::Number(25.0),
            ],
        ),
        PagesTableFormulaCachedValue::Number(125.0),
    )?;

    let anchor = editor.body_text()?.encode_utf16().count();
    let copy = editor.duplicate_table(source.model_object_id, anchor)?;
    editor.set_table_cell(
        copy.model_object_id,
        1,
        0,
        PagesCellValue::Text("Lodging".to_owned()),
    )?;
    editor.save(&output)?;
    println!("created {output} with {} and {}", source.name, copy.name);
    Ok(())
}
