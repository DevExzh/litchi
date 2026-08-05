//! Create overlapping Numbers shapes and move one to the front without a template.

use litchi_iwa::DrawableLayerMove;
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeFill};
use litchi_iwa_common::shape::path::Preset;

const OVERLAP_POSITION: DrawablePoint = DrawablePoint { x: 420.0, y: 300.0 };
const OVERLAP_SIZE: DrawableSize = DrawableSize {
    width: 300.0,
    height: 150.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_numbers_stacked_shapes <output.numbers>")?;
    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Layered Shapes")
        .table_name("Data")
        .build()?;
    let sheet_id = editor.sheets()?[0].object_id;
    let first = editor.add_sheet_shape_with_fill(
        sheet_id,
        "Moved to Front",
        OVERLAP_POSITION,
        OVERLAP_SIZE,
        Preset::Rectangle,
        solid_color(0.92, 0.18, 0.16)?,
    )?;
    let second = editor.add_sheet_shape_with_fill(
        sheet_id,
        "Originally Front",
        OVERLAP_POSITION,
        OVERLAP_SIZE,
        Preset::Ellipse,
        solid_color(0.12, 0.38, 0.92)?,
    )?;
    editor.move_sheet_drawable(
        sheet_id,
        first.drawable_object_id,
        DrawableLayerMove::ToFront,
    )?;
    editor.save(output)?;
    println!(
        "moved Numbers drawable {} ahead of {}",
        first.drawable_object_id, second.drawable_object_id
    );
    Ok(())
}

fn solid_color(red: f32, green: f32, blue: f32) -> litchi_iwa::Result<ShapeFill> {
    Ok(ShapeFill::Solid(RgbaColor::new(
        red,
        green,
        blue,
        1.0,
        RgbColorSpace::Srgb,
    )?))
}
