//! Concise typed PresentationML slide demo.
//!
//! Run with: cargo run --example pptx_simple_demo --features ooxml

use litchi_pptx::{MutablePresentation, Package};

const EMU_PER_INCH: i64 = 914_400;
const TEXT_WIDTH: i64 = 7_315_200;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating typed PPTX demonstration...\n");

    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        presentation.set_widescreen_slide_size();
        add_title_slide(presentation)?;
        add_shapes_slide(presentation)?;
    }

    let output_path = "pptx_simple_demo.pptx";
    println!("\nSaving to {}...", output_path);
    package.save(output_path)?;
    println!("✓ Done!");

    Ok(())
}

fn add_title_slide(
    presentation: &mut MutablePresentation,
) -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating Slide 1: Title");
    let slide = presentation.add_slide()?;
    slide.set_title("Simple Typed Demo");
    slide
        .add_text_box(
            "A small PresentationML deck authored with typed slide and shape models.",
            EMU_PER_INCH,
            3 * EMU_PER_INCH,
            TEXT_WIDTH,
            EMU_PER_INCH,
        )
        .font("Aptos")
        .font_size(24.0);
    Ok(())
}

fn add_shapes_slide(
    presentation: &mut MutablePresentation,
) -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating Slide 2: Shapes");
    let slide = presentation.add_slide()?;
    slide.set_title("Typed Shapes");
    slide.add_text_box(
        "Rectangle and ellipse shapes use the current typed writer API.",
        EMU_PER_INCH,
        EMU_PER_INCH,
        TEXT_WIDTH,
        EMU_PER_INCH,
    );
    slide.add_rectangle(
        EMU_PER_INCH,
        2 * EMU_PER_INCH,
        2 * EMU_PER_INCH,
        1_500_000,
        Some("FF6B6B".to_owned()),
    );
    slide.add_ellipse(
        4 * EMU_PER_INCH,
        2 * EMU_PER_INCH,
        2 * EMU_PER_INCH,
        1_500_000,
        Some("4ECDC4".to_owned()),
    );
    Ok(())
}
