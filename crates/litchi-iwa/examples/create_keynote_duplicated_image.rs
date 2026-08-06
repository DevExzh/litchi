//! Create a Keynote presentation with an image and its native-style duplicate.

use std::env;
use std::fs;
use std::path::Path;

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa_common::shape::geometry::{Point, Size};
use litchi_keynote::slide::image::Options as ImageOptions;

const IMAGE_POSITION: Point = Point { x: 704.0, y: 284.0 };
const IMAGE_SIZE: Size = Size {
    width: 512.0,
    height: 512.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_keynote_duplicated_image <output.key> <image>")?;
    let image_path = arguments
        .next()
        .ok_or("usage: create_keynote_duplicated_image <output.key> <image>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    let preferred_filename = Path::new(&image_path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or("image path must end in a UTF-8 file name")?;
    let image = fs::read(&image_path)?;

    let mut editor = KeynoteDocumentBuilder::new()
        .title("Created from scratch")
        .subtitle("Native-style image duplication from typed IWA objects")
        .build()?;
    let source = editor.add_slide_image(
        0,
        preferred_filename,
        &image,
        ImageOptions::new(IMAGE_POSITION, IMAGE_SIZE)?,
    )?;
    let duplicate = editor.duplicate_slide_image(0, source.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "created Keynote images {} and {} sharing data {}",
        source.drawable_object_id, duplicate.drawable_object_id, source.image_data_identifier
    );
    Ok(())
}
