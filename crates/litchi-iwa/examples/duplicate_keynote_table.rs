use std::env;

use litchi_iwa::keynote::{self};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_keynote::slide::table::formula::{FormulaCachedValue, FormulaExpression};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = env::args()
        .nth(1)
        .unwrap_or_else(|| "keynote-duplicated-table.key".to_owned());
    let mut editor = keynote::Builder::new().title("Budget scenarios").build()?;
    let source = editor.add_slide_table(
        0,
        "Budget",
        3,
        2,
        DrawablePoint { x: 360.0, y: 360.0 },
        DrawableSize {
            width: 900.0,
            height: 300.0,
        },
    )?;
    editor.set_slide_table_cells(
        0,
        source.model_object_id,
        [
            keynote::TableCellUpdate::new(
                0,
                0,
                keynote::TableCellValue::Text("Category".to_owned()),
            ),
            keynote::TableCellUpdate::new(0, 1, keynote::TableCellValue::Text("Cost".to_owned())),
            keynote::TableCellUpdate::new(1, 0, keynote::TableCellValue::Text("Travel".to_owned())),
            keynote::TableCellUpdate::new(1, 1, keynote::TableCellValue::number(125.0)?),
        ],
    )?;
    editor.set_slide_table_formula(
        0,
        source.model_object_id,
        2,
        1,
        FormulaExpression::function(
            "SUM",
            [
                FormulaExpression::Number(100.0),
                FormulaExpression::Number(25.0),
            ],
        ),
        FormulaCachedValue::number(125.0)?,
    )?;

    let copy = editor.duplicate_slide_table(0, source.drawable_object_id)?;
    editor.set_slide_table_cell(
        0,
        copy.model_object_id,
        1,
        0,
        keynote::TableCellValue::Text("Lodging".to_owned()),
    )?;
    editor.save(&output)?;
    println!("created {output} with {} and {}", source.name, copy.name);
    Ok(())
}
