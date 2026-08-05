use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa::text::{TextNumberAttachmentKind, TextNumberAttachmentSettings, TextPosition};
use litchi_iwa_common::comment::DrawableId;

const BODY: &str = "Body";
const BODY_ANCHOR: usize = 4;
const TEXT_BOX_PREFIX: &str = "Text box page ";
const TEXT_BOX_POSITION: DrawablePoint = DrawablePoint { x: 80.0, y: 160.0 };
const TEXT_BOX_SIZE: DrawableSize = DrawableSize {
    width: 280.0,
    height: 80.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_pages_text_box_number_attachment <output.pages>")?;
    let mut editor = PagesEditor::create_with_text(BODY)?;
    let text_box = editor.add_text_box(
        BODY_ANCHOR,
        TEXT_BOX_PREFIX,
        TEXT_BOX_POSITION,
        TEXT_BOX_SIZE,
    )?;
    let drawable_object_id = DrawableId::from_raw(text_box.drawable_object_id)?;
    editor.insert_text_box_number_attachment(
        drawable_object_id,
        TextPosition::from_utf16_index(TEXT_BOX_PREFIX.encode_utf16().count())?,
        TextNumberAttachmentSettings::new(TextNumberAttachmentKind::PageNumber),
    )?;
    editor.save(output)?;
    Ok(())
}
