use litchi_iwa::pages::PagesEditor;
use litchi_iwa::text::{TextNumberAttachmentKind, TextNumberAttachmentSettings, TextPosition};

const BODY: &str = "Page  of ";
const PAGE_NUMBER_POSITION: usize = 5;
const PAGE_COUNT_POSITION: usize = 10;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_pages_number_attachments <output.pages>")?;
    let mut editor = PagesEditor::create_with_text(BODY)?;
    editor.insert_body_number_attachment(
        TextPosition::from_utf16_index(PAGE_NUMBER_POSITION)?,
        TextNumberAttachmentSettings::new(TextNumberAttachmentKind::PageNumber),
    )?;
    editor.insert_body_number_attachment(
        TextPosition::from_utf16_index(PAGE_COUNT_POSITION)?,
        TextNumberAttachmentSettings::new(TextNumberAttachmentKind::PageCount),
    )?;
    editor.save(output)?;
    Ok(())
}
