//! Create a Pages document and editable preset shape without an input package.

use std::{env, fs, path::Path};

use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{
    DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeEffects, ShapeFill, ShapeGradient,
    ShapeGradientAngle, ShapeImageFillTechnique, ShapeOpacity, ShapePreset, ShapeReflection,
    ShapeReflectionOpacity,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_pages_shape <output.pages> [text] [fill-image]")?;
    let text = arguments
        .next()
        .unwrap_or_else(|| "Built from typed IWA objects".to_owned());
    let fill_image = arguments.next();
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let body = "Pages shape created entirely by litchi-iwa";
    let mut editor = PagesEditor::create_with_text(body)?;
    let created = editor.add_body_shape_with_fill(
        body.encode_utf16().count(),
        &text,
        DrawablePoint { x: 180.0, y: 240.0 },
        DrawableSize {
            width: 300.0,
            height: 150.0,
        },
        ShapePreset::RightArrow,
        ShapeFill::Gradient(ShapeGradient::linear(
            RgbaColor::new(0.88, 0.18, 0.12, 1.0, RgbColorSpace::DisplayP3)?,
            RgbaColor::new(0.98, 0.65, 0.08, 1.0, RgbColorSpace::DisplayP3)?,
            ShapeGradientAngle::from_degrees(45.0)?,
        )),
    )?;
    if let Some(path) = fill_image {
        let filename = preferred_filename(&path)?;
        editor.set_body_shape_image_fill(
            created.drawable_object_id,
            filename,
            &fs::read(&path)?,
            ShapeImageFillTechnique::ScaleToFill,
            None,
        )?;
    }
    editor.set_body_shape_effects(
        created.drawable_object_id,
        ShapeEffects::new(
            ShapeOpacity::new(0.72)?,
            ShapeReflection::Enabled(ShapeReflectionOpacity::new(0.35)?),
        ),
    )?;
    editor.save(output)?;
    println!(
        "created Pages {:?} {:?} {} with storage {} at UTF-16 anchor {}",
        created.kind,
        created.preset,
        created.drawable_object_id,
        created.storage.object_id,
        created.anchor_character_index
    );
    Ok(())
}

fn preferred_filename(path: &str) -> Result<&str, Box<dyn std::error::Error>> {
    Path::new(path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| "fill image path has no UTF-8 filename".into())
}
