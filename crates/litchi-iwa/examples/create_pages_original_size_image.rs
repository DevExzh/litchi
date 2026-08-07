//! Create a Pages image and restore its typed original dimensions from scratch.

use std::fs;
use std::path::Path;

use litchi_iwa::pages::{PagesEditor, PagesImageOptions};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::comment::DrawableId;

const BODY_TEXT: &str = "This image was restored to its original size by litchi-iwa.";
const IMAGE_POSITION: DrawablePoint = DrawablePoint { x: 64.0, y: 128.0 };
const DISPLAYED_IMAGE_SIZE: DrawableSize = DrawableSize {
    width: 240.0,
    height: 240.0,
};
const ORIGINAL_IMAGE_SIZE: DrawableSize = DrawableSize {
    width: 512.0,
    height: 512.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_pages_original_size_image <output.pages> <image>")?;
    let image_path = arguments
        .next()
        .ok_or("usage: create_pages_original_size_image <output.pages> <image>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let image_filename = filename(&image_path)?;
    let image = fs::read(&image_path)?;

    let mut editor = PagesEditor::create_with_text(BODY_TEXT)?;
    let created = editor.add_body_image(
        BODY_TEXT.encode_utf16().count(),
        image_filename,
        &image,
        PagesImageOptions::new(IMAGE_POSITION, DISPLAYED_IMAGE_SIZE)
            .with_natural_size(ORIGINAL_IMAGE_SIZE),
    )?;
    editor.restore_body_image_original_size(DrawableId::from_raw(created.drawable_object_id)?)?;
    editor.save(output)?;
    println!(
        "created Pages image {} at its original size",
        created.drawable_object_id
    );
    Ok(())
}

fn filename(path: &str) -> Result<&str, Box<dyn std::error::Error>> {
    Path::new(path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| "image path must end in a UTF-8 file name".into())
}
