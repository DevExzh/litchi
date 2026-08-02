//! Create a Pages document with an image and its native-style duplicate.

use std::fs;
use std::path::Path;

use litchi_iwa::pages::{PagesEditor, PagesImageOptions};
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
        .ok_or("usage: create_pages_duplicated_image <output.pages> <image> [body text]")?;
    let image_path = arguments
        .next()
        .ok_or("usage: create_pages_duplicated_image <output.pages> <image> [body text]")?;
    let body_text = arguments
        .next()
        .unwrap_or_else(|| "Created from scratch with litchi-iwa".to_owned());
    let image = fs::read(&image_path)?;
    let filename = Path::new(&image_path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or("image path has no UTF-8 filename")?;

    let mut editor = PagesEditor::create_with_text(body_text)?;
    let source = editor.add_body_image(
        editor.body_text()?.encode_utf16().count(),
        filename,
        &image,
        PagesImageOptions::new(IMAGE_POSITION, IMAGE_SIZE),
    )?;
    let duplicate = editor.duplicate_body_image(
        source.drawable_object_id,
        editor.body_text()?.encode_utf16().count(),
    )?;
    editor.save(output)?;
    println!(
        "created Pages images {} and {} sharing data {}",
        source.drawable_object_id, duplicate.drawable_object_id, source.image_data_identifier
    );
    Ok(())
}
