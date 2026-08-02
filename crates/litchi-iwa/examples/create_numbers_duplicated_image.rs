//! Create a Numbers spreadsheet with an image and its native-style duplicate.

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
        .ok_or("usage: create_numbers_duplicated_image <output.numbers> <image>")?;
    let image_path = arguments
        .next()
        .ok_or("usage: create_numbers_duplicated_image <output.numbers> <image>")?;
    let preferred_filename = Path::new(&image_path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or("image path must end in a UTF-8 file name")?;
    let image = fs::read(&image_path)?;

    let mut editor = NumbersDocumentBuilder::new().sheet_name("Media").build()?;
    let sheet_id = editor.sheets()?[0].object_id;
    let source = editor.add_sheet_image(
        sheet_id,
        preferred_filename,
        &image,
        NumbersSheetImageOptions::new(IMAGE_POSITION, IMAGE_SIZE),
    )?;
    let duplicate = editor.duplicate_sheet_image(sheet_id, source.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "created Numbers images {} and {} sharing data {}",
        source.drawable_object_id, duplicate.drawable_object_id, source.image_data_identifier
    );
    Ok(())
}
