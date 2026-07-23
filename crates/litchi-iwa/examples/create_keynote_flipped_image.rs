//! Create a Keynote image and apply a native horizontal Arrange flip from scratch.

use std::fs;
use std::path::Path;

use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteSlideImageOptions};
use litchi_iwa::shapes::{DrawableFlipAxis, DrawablePoint, DrawableSize};

const IMAGE_POSITION: DrawablePoint = DrawablePoint { x: 704.0, y: 284.0 };
const IMAGE_SIZE: DrawableSize = DrawableSize {
    width: 512.0,
    height: 512.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_keynote_flipped_image <output.key> <image>")?;
    let image_path = arguments
        .next()
        .ok_or("usage: create_keynote_flipped_image <output.key> <image>")?;
    let preferred_filename = Path::new(&image_path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or("image path must end in a UTF-8 file name")?;
    let image = fs::read(&image_path)?;

    let mut editor = KeynoteDocumentBuilder::new()
        .title("Flipped Image")
        .subtitle("Typed native Arrange flip")
        .build()?;
    let created = editor.add_slide_image(
        0,
        preferred_filename,
        &image,
        KeynoteSlideImageOptions::new(IMAGE_POSITION, IMAGE_SIZE),
    )?;
    editor.flip_slide_image(0, created.drawable_object_id, DrawableFlipAxis::Horizontal)?;
    editor.save(output)?;
    println!(
        "created horizontally flipped Keynote image {}",
        created.drawable_object_id
    );
    Ok(())
}
