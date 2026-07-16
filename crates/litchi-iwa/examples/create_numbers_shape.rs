//! Create a Numbers spreadsheet and editable preset shape without an input package.

use std::{env, fs, path::Path};

use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::shapes::{
    DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeContactShadow, ShapeEffects,
    ShapeFill, ShapeGradient, ShapeGradientAngle, ShapeImageFillTechnique, ShapeOpacity,
    ShapePreset, ShapeReflection, ShapeReflectionOpacity, ShapeShadow, ShapeShadowAppearance,
    ShapeShadowBlurRadius, ShapeShadowOffset, ShapeShadowOpacity, ShapeShadowPerspective,
    ShapeTextAutoSize, ShapeTextInset, ShapeTextInsets, ShapeTextLayout,
    ShapeTextVerticalAlignment,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_numbers_shape <output.numbers> [text] [fill-image]")?;
    let text = arguments
        .next()
        .unwrap_or_else(|| "Built from typed IWA objects".to_owned());
    let fill_image = arguments.next();
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Scratch Shape")
        .table_name("Scratch Table")
        .build()?;
    let sheet_id = editor.sheets()?[0].object_id;
    let created = editor.add_sheet_shape_with_fill(
        sheet_id,
        &text,
        DrawablePoint { x: 420.0, y: 300.0 },
        DrawableSize {
            width: 300.0,
            height: 150.0,
        },
        ShapePreset::RightArrow,
        ShapeFill::Gradient(ShapeGradient::linear(
            RgbaColor::new(0.08, 0.42, 0.9, 1.0, RgbColorSpace::Srgb)?,
            RgbaColor::new(0.1, 0.85, 0.78, 1.0, RgbColorSpace::Srgb)?,
            ShapeGradientAngle::from_degrees(0.0)?,
        )),
    )?;
    if let Some(path) = fill_image {
        let filename = preferred_filename(&path)?;
        editor.set_sheet_shape_image_fill(
            sheet_id,
            created.drawable_object_id,
            filename,
            &fs::read(&path)?,
            ShapeImageFillTechnique::ScaleToFit,
            None,
        )?;
    }
    editor.set_sheet_shape_effects(
        sheet_id,
        created.drawable_object_id,
        ShapeEffects::new(
            ShapeOpacity::new(0.84)?,
            ShapeReflection::Enabled(ShapeReflectionOpacity::new(0.65)?),
        ),
    )?;
    editor.set_sheet_shape_shadow(
        sheet_id,
        created.drawable_object_id,
        ShapeShadow::Contact(ShapeContactShadow::new(
            ShapeShadowAppearance::new(
                RgbaColor::black(),
                ShapeShadowBlurRadius::from_points(18)?,
                ShapeShadowOffset::from_points(6.0)?,
                ShapeShadowOpacity::new(0.58)?,
            ),
            ShapeShadowPerspective::from_degrees(23.0)?,
        )),
    )?;
    editor.set_sheet_shape_text_layout(
        sheet_id,
        created.drawable_object_id,
        ShapeTextLayout::new(
            ShapeTextVerticalAlignment::Bottom,
            ShapeTextInsets::uniform(ShapeTextInset::from_points(9.0)?),
            ShapeTextAutoSize::Fixed,
        ),
    )?;
    editor.save(output)?;
    println!(
        "created Numbers {:?} {:?} {} with storage {} on sheet {}",
        created.kind,
        created.preset,
        created.drawable_object_id,
        created.storage.object_id,
        sheet_id
    );
    Ok(())
}

fn preferred_filename(path: &str) -> Result<&str, Box<dyn std::error::Error>> {
    Path::new(path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| "fill image path has no UTF-8 filename".into())
}
