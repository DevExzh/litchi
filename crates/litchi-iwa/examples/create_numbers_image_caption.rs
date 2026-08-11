//! Create a Numbers spreadsheet with an image title and caption from scratch.

use std::fs;
use std::path::Path;

use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersSheetImageOptions};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

const IMAGE_POSITION: DrawablePoint = DrawablePoint { x: 420.0, y: 180.0 };
const IMAGE_SIZE: DrawableSize = DrawableSize {
    width: 320.0,
    height: 240.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_numbers_image_caption <output.numbers> <image> <title> <caption>")?;
    let image_path = arguments
        .next()
        .ok_or("usage: create_numbers_image_caption <output.numbers> <image> <title> <caption>")?;
    let title = arguments
        .next()
        .ok_or("usage: create_numbers_image_caption <output.numbers> <image> <title> <caption>")?;
    let caption = arguments
        .next()
        .ok_or("usage: create_numbers_image_caption <output.numbers> <image> <title> <caption>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let image = fs::read(&image_path)?;
    let filename = Path::new(&image_path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or("image path has no UTF-8 filename")?;
    let mut editor = NumbersDocumentBuilder::new().sheet_name("Media").build()?;
    let sheet_id = editor.sheets()?[0].id();
    let created = editor.add_sheet_image(
        sheet_id,
        filename,
        &image,
        NumbersSheetImageOptions::new(IMAGE_POSITION, IMAGE_SIZE),
    )?;
    editor.set_sheet_image_title(sheet_id, created.drawable_object_id, &title)?;
    editor.set_sheet_image_caption(sheet_id, created.drawable_object_id, &caption)?;
    editor.save(output)?;
    Ok(())
}
