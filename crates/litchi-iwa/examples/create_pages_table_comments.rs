//! Create a Pages document with a native table-comment thread from scratch.

use std::path::PathBuf;

use litchi_iwa::pages::PagesDocumentBuilder;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_pages_table_comments <output.pages>")?,
    );
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Native table-comment creation\n")
        .body_table("Review", 3, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.set_table_cell_comment(table_id, 1, 1, "Pages root comment")?;
    editor.add_table_cell_comment_reply(table_id, 1, 1, "Pages direct reply")?;
    editor.save(output)?;
    Ok(())
}
