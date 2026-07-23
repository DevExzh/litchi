//! Create a Numbers image and restore its typed original dimensions from scratch.

use std::fs;
use std::path::Path;

use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersSheetImageOptions};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

const IMAGE_POSITION: DrawablePoint = DrawablePoint { x: 420.0, y: 180.0 };
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
        .ok_or("usage: create_numbers_original_size_image <output.numbers> <image>")?;
    let image_path = arguments
        .next()
        .ok_or("usage: create_numbers_original_size_image <output.numbers> <image>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let image_filename = filename(&image_path)?;
    let image = fs::read(&image_path)?;

    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Original Size Image")
        .build()?;
    let sheet_id = editor.sheets()?[0].object_id;
    let created = editor.add_sheet_image(
        sheet_id,
        image_filename,
        &image,
        NumbersSheetImageOptions::new(IMAGE_POSITION, DISPLAYED_IMAGE_SIZE)
            .with_natural_size(ORIGINAL_IMAGE_SIZE),
    )?;
    editor.restore_sheet_image_original_size(sheet_id, created.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "created Numbers image {} at its original size",
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
