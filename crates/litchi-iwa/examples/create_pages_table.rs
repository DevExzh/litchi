use std::env;

use litchi_iwa::pages::{PagesCellValue, PagesDocumentBuilder};

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
