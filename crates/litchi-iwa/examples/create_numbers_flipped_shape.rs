//! Create a Numbers shape and apply a native horizontal Arrange flip from scratch.

use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::shapes::{DrawableFlipAxis, DrawablePoint, DrawableSize};
use litchi_iwa_common::shape::path::Preset;

const ARROW_POSITION: DrawablePoint = DrawablePoint { x: 420.0, y: 300.0 };
const ARROW_SIZE: DrawableSize = DrawableSize {
    width: 300.0,
    height: 150.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_numbers_flipped_shape <output.numbers>")?;
    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Flipped Shape")
        .table_name("Data")
        .build()?;
    let sheet_id = editor.sheets()?[0].object_id;
    let arrow = editor.add_sheet_shape(
        sheet_id,
        "Horizontally Flipped",
        ARROW_POSITION,
        ARROW_SIZE,
        Preset::RightArrow,
    )?;
    editor.flip_sheet_shape(
        sheet_id,
        arrow.drawable_object_id,
        DrawableFlipAxis::Horizontal,
    )?;
    editor.save(output)?;
    println!(
        "created horizontally flipped Numbers arrow {}",
        arrow.drawable_object_id
    );
    Ok(())
}
