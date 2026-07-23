use std::fs;

use litchi_iwa::pages::{PagesEditor, PagesImageOptions};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa::{ImageAdjustment, ImageAdjustments, ImageEnhancement};

const IMAGE_POSITION: DrawablePoint = DrawablePoint { x: 96.0, y: 144.0 };
const IMAGE_SIZE: DrawableSize = DrawableSize {
    width: 300.0,
    height: 225.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_pages_image <output.pages> <image> [body text]")?;
    let image_path = arguments
        .next()
        .ok_or("usage: create_pages_image <output.pages> <image> [body text]")?;
    let body_text = arguments
        .next()
        .unwrap_or_else(|| "Created from scratch with litchi-iwa".to_owned());
    let anchor = body_text.encode_utf16().count();
    let image = fs::read(&image_path)?;
    let filename = std::path::Path::new(&image_path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or("image path has no UTF-8 filename")?;

    let mut editor = PagesEditor::create_with_text(body_text)?;
    let created = editor.add_body_image(
        anchor,
        filename,
        &image,
        PagesImageOptions::new(IMAGE_POSITION, IMAGE_SIZE),
    )?;
    let mut properties = editor.body_image_properties(created.drawable_object_id)?;
    properties.accessibility_description = Some(format!("Embedded image: {filename}"));
    editor.set_body_image_properties(created.drawable_object_id, properties)?;
    editor.set_body_image_adjustments(
        created.drawable_object_id,
        ImageAdjustments::default()
            .with_exposure(Some(ImageAdjustment::new(0.25)?))
            .with_saturation(Some(ImageAdjustment::new(-0.5)?))
            .with_enhancement(Some(ImageEnhancement::Enabled)),
    )?;
    editor.save(output)?;
    Ok(())
}
