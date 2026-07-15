use std::fs;

use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

const IMAGE_POSITION: DrawablePoint = DrawablePoint { x: 96.0, y: 144.0 };
const IMAGE_SIZE: DrawableSize = DrawableSize {
    width: 300.0,
    height: 225.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_pages_image <output.pages> <image> [body text]")?;
    let image_path = arguments
        .next()
        .ok_or("usage: create_pages_image <output.pages> <image> [body text]")?;
    let body_text = arguments
        .next()
        .unwrap_or_else(|| "Created from scratch with litchi-iwa".to_owned());
    let anchor = body_text.encode_utf16().count();
    let image = fs::read(&image_path)?;
    let filename = std::path::Path::new(&image_path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or("image path has no UTF-8 filename")?;

    let mut editor = PagesEditor::create_with_text(body_text)?;
    editor.add_body_image(anchor, filename, &image, IMAGE_POSITION, IMAGE_SIZE)?;
    editor.save(output)?;
    Ok(())
}
