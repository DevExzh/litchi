use litchi_iwa::pages::PagesEditor;
use litchi_iwa::text::{TextBookmarkName, TextBookmarkSettings, TextRange};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_pages_bookmark <output.pages>")?;
    let mut pages = PagesEditor::create_with_text("Methods and results")?;
    let settings = TextBookmarkSettings::new().with_name(TextBookmarkName::new("Methods")?);
    pages.add_body_bookmark(TextRange::from_utf16_indexes(0, 7)?, settings)?;
    pages.save(output)?;
    Ok(())
}
