//! Create a Keynote image and restore its typed original dimensions from scratch.

use std::fs;
use std::path::Path;

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa_common::shape::geometry::{Point, Size};
use litchi_keynote::slide::image::Options as ImageOptions;

const IMAGE_POSITION: Point = Point { x: 720.0, y: 405.0 };
const DISPLAYED_IMAGE_SIZE: Size = Size {
    width: 240.0,
    height: 240.0,
};
const ORIGINAL_IMAGE_SIZE: Size = Size {
    width: 512.0,
    height: 512.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_keynote_original_size_image <output.key> <image>")?;
    let image_path = arguments
        .next()
        .ok_or("usage: create_keynote_original_size_image <output.key> <image>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let image_filename = filename(&image_path)?;
    let image = fs::read(&image_path)?;

    let mut editor = KeynoteDocumentBuilder::new()
        .title("Original Size Image")
        .subtitle("Typed native image dimensions")
        .build()?;
    let created = editor.add_slide_image(
        0,
        image_filename,
        &image,
        ImageOptions::new(IMAGE_POSITION, DISPLAYED_IMAGE_SIZE)?
            .with_natural_size(ORIGINAL_IMAGE_SIZE)?,
    )?;
    editor.restore_slide_image_original_size(0, created.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "created Keynote image {} at its original size",
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
