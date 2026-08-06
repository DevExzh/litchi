//! Create a Keynote presentation with an image title and caption from scratch.

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
    let mut arguments = std::env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_keynote_image_caption <output.key> <image> <title> <caption>")?;
    let image_path = arguments
        .next()
        .ok_or("usage: create_keynote_image_caption <output.key> <image> <title> <caption>")?;
    let title = arguments
        .next()
        .ok_or("usage: create_keynote_image_caption <output.key> <image> <title> <caption>")?;
    let caption = arguments
        .next()
        .ok_or("usage: create_keynote_image_caption <output.key> <image> <title> <caption>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let image = fs::read(&image_path)?;
    let filename = Path::new(&image_path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or("image path has no UTF-8 filename")?;
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Created from scratch")
        .subtitle("Image labels built from typed IWA objects")
        .build()?;
    let created = editor.add_slide_image(
        0,
        filename,
        &image,
        ImageOptions::new(IMAGE_POSITION, IMAGE_SIZE)?,
    )?;
    editor.set_slide_image_title(0, created.drawable_object_id, &title)?;
    editor.set_slide_image_caption(0, created.drawable_object_id, &caption)?;
    editor.save(output)?;
    Ok(())
}
