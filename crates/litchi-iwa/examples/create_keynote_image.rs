//! Create a Keynote presentation with an image and no input package.

use std::env;
use std::fs;
use std::path::Path;

use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteSlideImageOptions};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::shape::image::{ImageAdjustment, ImageAdjustments, ImageEnhancement};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_keynote_image <output.key> <image>")?;
    let image_path = arguments
        .next()
        .ok_or("usage: create_keynote_image <output.key> <image>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    let preferred_filename = Path::new(&image_path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or("image path must end in a UTF-8 file name")?;
    let image = fs::read(&image_path)?;

    let mut editor = KeynoteDocumentBuilder::new()
        .title("Created from scratch")
        .subtitle("Image built from typed IWA objects")
        .build()?;
    let created = editor.add_slide_image(
        0,
        preferred_filename,
        &image,
        KeynoteSlideImageOptions::new(
            DrawablePoint { x: 704.0, y: 284.0 },
            DrawableSize {
                width: 512.0,
                height: 512.0,
            },
        ),
    )?;
    let mut properties = editor.slide_image_properties(0, created.drawable_object_id)?;
    properties.accessibility_description = Some(format!("Embedded image: {preferred_filename}"));
    editor.set_slide_image_properties(0, created.drawable_object_id, properties)?;
    editor.set_slide_image_adjustments(
        0,
        created.drawable_object_id,
        ImageAdjustments::default()
            .with_exposure(Some(ImageAdjustment::new(0.25)?))
            .with_saturation(Some(ImageAdjustment::new(-0.5)?))
            .with_enhancement(Some(ImageEnhancement::Enabled)),
    )?;
    editor.save(output)?;
    println!(
        "created Keynote image {} backed by data {}",
        created.drawable_object_id, created.image_data_identifier
    );
    Ok(())
}
