//! Create a Numbers image and apply a native horizontal Arrange flip from scratch.

use std::fs;
use std::path::Path;

use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersSheetImageOptions};
use litchi_iwa::shapes::{DrawableFlipAxis, DrawablePoint, DrawableSize};

const IMAGE_POSITION: DrawablePoint = DrawablePoint { x: 420.0, y: 180.0 };
const IMAGE_SIZE: DrawableSize = DrawableSize {
    width: 320.0,
    height: 240.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_numbers_flipped_image <output.numbers> <image>")?;
    let image_path = arguments
        .next()
        .ok_or("usage: create_numbers_flipped_image <output.numbers> <image>")?;
    let preferred_filename = Path::new(&image_path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or("image path must end in a UTF-8 file name")?;
    let image = fs::read(&image_path)?;

    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Flipped Image")
        .build()?;
    let sheet_id = editor.sheets()?[0].id();
    let created = editor.add_sheet_image(
        sheet_id,
        preferred_filename,
        &image,
        NumbersSheetImageOptions::new(IMAGE_POSITION, IMAGE_SIZE),
    )?;
    editor.flip_sheet_image(
        sheet_id,
        created.drawable_object_id,
        DrawableFlipAxis::Horizontal,
    )?;
    editor.save(output)?;
    println!(
        "created horizontally flipped Numbers image {}",
        created.drawable_object_id
    );
    Ok(())
}
