//! Create a Pages image and apply a native horizontal Arrange flip from scratch.

use std::fs;
use std::path::Path;

use litchi_iwa::pages::{PagesEditor, PagesImageOptions};
use litchi_iwa::shapes::{DrawableFlipAxis, DrawablePoint, DrawableSize};

const IMAGE_POSITION: DrawablePoint = DrawablePoint { x: 96.0, y: 144.0 };
const IMAGE_SIZE: DrawableSize = DrawableSize {
    width: 300.0,
    height: 225.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_pages_flipped_image <output.pages> <image>")?;
    let image_path = arguments
        .next()
        .ok_or("usage: create_pages_flipped_image <output.pages> <image>")?;
    let preferred_filename = Path::new(&image_path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or("image path must end in a UTF-8 file name")?;
    let image = fs::read(&image_path)?;
    let body = "This mirrored image was created entirely by litchi-iwa.";

    let mut editor = PagesEditor::create_with_text(body)?;
    let created = editor.add_body_image(
        body.encode_utf16().count(),
        preferred_filename,
        &image,
        PagesImageOptions::new(IMAGE_POSITION, IMAGE_SIZE),
    )?;
    let drawable_id = litchi_iwa::comments::DrawableObjectId::new(created.drawable_object_id)
        .ok_or("created image has an invalid drawable identifier")?;
    editor.flip_body_image(drawable_id, DrawableFlipAxis::Horizontal)?;
    editor.save(output)?;
    println!(
        "created horizontally flipped Pages image {}",
        created.drawable_object_id
    );
    Ok(())
}
