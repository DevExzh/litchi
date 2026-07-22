//! Create and physically sort a plain-text Pages table without an input document.

use std::env;

use litchi_iwa::pages::{
    PagesCellValue, PagesDocumentBuilder, PagesTableCellUpdate, PagesTableHeaderCount,
    PagesTableHeaderSettings, PagesTableSortColumnIndex, PagesTableSortDirection,
    PagesTableSortOrder, PagesTableSortRule,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = env::args()
        .nth(1)
        .ok_or("usage: create_pages_sorted_table <output.pages>")?;
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Cities sorted by name\n")
        .body_table("Cities", 5, 2)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.set_table_header_settings(
        table_id,
        PagesTableHeaderSettings {
            header_rows: Some(PagesTableHeaderCount::ONE),
            ..Default::default()
        },
    )?;
    editor.set_table_cells(
        table_id,
        [
            PagesTableCellUpdate::new(0, 0, PagesCellValue::Text("Name".to_owned())),
            PagesTableCellUpdate::new(0, 1, PagesCellValue::Text("Marker".to_owned())),
            PagesTableCellUpdate::new(1, 0, PagesCellValue::Text("zebra".to_owned())),
            PagesTableCellUpdate::new(1, 1, PagesCellValue::Text("last".to_owned())),
            PagesTableCellUpdate::new(2, 0, PagesCellValue::Text("apple".to_owned())),
            PagesTableCellUpdate::new(2, 1, PagesCellValue::Text("first apple".to_owned())),
            PagesTableCellUpdate::new(3, 0, PagesCellValue::Text("banana".to_owned())),
            PagesTableCellUpdate::new(3, 1, PagesCellValue::Text("middle".to_owned())),
            PagesTableCellUpdate::new(4, 0, PagesCellValue::Text("apple".to_owned())),
            PagesTableCellUpdate::new(4, 1, PagesCellValue::Text("second apple".to_owned())),
        ],
    )?;
    editor.set_table_sort_order(
        table_id,
        PagesTableSortOrder::new([PagesTableSortRule::new(
            PagesTableSortColumnIndex::new(0)?,
            PagesTableSortDirection::Ascending,
        )])?,
    )?;
    if !editor.apply_table_sort_order(table_id)? {
        return Err("expected the source table to be reordered".into());
    }
    editor.save(&output)?;
    println!("created {output}");
    Ok(())
}
