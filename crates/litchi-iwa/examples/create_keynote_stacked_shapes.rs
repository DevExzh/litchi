//! Create overlapping Keynote shapes and move one to the front without a template.

use litchi_iwa::DrawableLayerMove;
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeFill};
use litchi_iwa_common::shape::path::Preset;

const OVERLAP_POSITION: DrawablePoint = DrawablePoint { x: 720.0, y: 660.0 };
const OVERLAP_SIZE: DrawableSize = DrawableSize {
    width: 480.0,
    height: 240.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_keynote_stacked_shapes <output.key>")?;
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Layered Shapes")
        .subtitle("Typed Arrange operations")
        .build()?;
    let first = editor.add_slide_shape_with_fill(
        0,
        "Moved to Front",
        OVERLAP_POSITION,
        OVERLAP_SIZE,
        Preset::Rectangle,
        solid_color(0.92, 0.18, 0.16)?,
    )?;
    let second = editor.add_slide_shape_with_fill(
        0,
        "Originally Front",
        OVERLAP_POSITION,
        OVERLAP_SIZE,
        Preset::Ellipse,
        solid_color(0.12, 0.38, 0.92)?,
    )?;
    editor.move_slide_drawable(0, first.drawable_object_id, DrawableLayerMove::ToFront)?;
    editor.save(output)?;
    println!(
        "moved Keynote drawable {} ahead of {}",
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
