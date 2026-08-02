use std::env;

use litchi_iwa::pages::PagesEditor;
use litchi_iwa::text::TextPosition;

const BODY: &str = "Alpha Beta";
const FOOTNOTE_POSITION: usize = 6;
const FOOTNOTE_TEXT: &str = "Created from scratch with litchi-iwa.";

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = env::args()
        .nth(1)
        .ok_or("usage: create_pages_footnotes <output.pages>")?;
    let mut pages = PagesEditor::create_with_text(BODY)?;
    let footnote = pages.insert_body_footnote(
        TextPosition::from_utf16_index(FOOTNOTE_POSITION)?,
        FOOTNOTE_TEXT,
    )?;
    pages.set_body_footnote_text(footnote.id, FOOTNOTE_TEXT)?;
    pages.save(output)?;
    Ok(())
}
