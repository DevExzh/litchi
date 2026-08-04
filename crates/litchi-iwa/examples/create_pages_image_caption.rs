//! Create a Pages document with an image title and caption from scratch.

use std::fs;
use std::path::Path;

use litchi_iwa::comments::DrawableObjectId;
use litchi_iwa::pages::{PagesEditor, PagesImageOptions};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

const IMAGE_POSITION: DrawablePoint = DrawablePoint { x: 96.0, y: 144.0 };
const IMAGE_SIZE: DrawableSize = DrawableSize {
    width: 300.0,
    height: 225.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let output = arguments.next().ok_or(
        "usage: create_pages_image_caption <output.pages> <image> <title> <caption> [body text]",
    )?;
    let image_path = arguments.next().ok_or(
        "usage: create_pages_image_caption <output.pages> <image> <title> <caption> [body text]",
    )?;
    let title = arguments.next().ok_or(
        "usage: create_pages_image_caption <output.pages> <image> <title> <caption> [body text]",
    )?;
    let caption = arguments.next().ok_or(
        "usage: create_pages_image_caption <output.pages> <image> <title> <caption> [body text]",
    )?;
    let body_text = arguments
        .next()
        .unwrap_or_else(|| "Created from scratch with litchi-iwa".to_owned());
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let image = fs::read(&image_path)?;
    let filename = Path::new(&image_path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or("image path has no UTF-8 filename")?;
    let mut editor = PagesEditor::create_with_text(&body_text)?;
    let created = editor.add_body_image(
        body_text.encode_utf16().count(),
        filename,
        &image,
        PagesImageOptions::new(IMAGE_POSITION, IMAGE_SIZE),
    )?;
    let drawable_object_id = DrawableObjectId::from_object_id(created.drawable_object_id)?;
    editor.set_body_image_title(drawable_object_id, &title)?;
    editor.set_body_image_caption(drawable_object_id, &caption)?;
    editor.save(output)?;
    Ok(())
}
