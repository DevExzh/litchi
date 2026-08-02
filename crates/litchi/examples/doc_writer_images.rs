//! Example demonstrating images in DOC files
//!
//! This example creates a DOC file with:
//! - Embedded JPEG images
//! - Embedded PNG images
//! - Images with different sizes
//! - Inline images in paragraphs
//!
//! Run with: cargo run --example doc_writer_images

use litchi::doc::writer::{CharacterFormatting, DocWriter, ParagraphFormatting};
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    println!("Creating DOC file with images...");

    // Create a new DOC writer
    let mut writer = DocWriter::new();

    // Add document title
    let title_char = CharacterFormatting {
        bold: Some(true),
        font_size: Some(28),
        ..Default::default()
    };
    let title_para = ParagraphFormatting {
        alignment: Some(1),
        ..Default::default()
    };
    writer.add_paragraph_with_format("Image Examples", title_char, title_para)?;
    writer.add_paragraph("")?;

    writer.add_paragraph(
        "This document demonstrates the current DOC writer workflow for image-oriented content.",
    )?;
    writer.add_paragraph("")?;

    // Section 1: JPEG image
    writer.add_paragraph("1. JPEG Image")?;
    writer.add_paragraph("Below is a JPEG image:")?;
    writer.add_paragraph("[Image 1 placeholder]")?;
    writer.add_paragraph("")?;

    // Section 2: PNG image
    writer.add_paragraph("2. PNG Image")?;
    writer.add_paragraph("Below is a PNG image with transparency:")?;
    writer.add_paragraph("[Image 2 placeholder]")?;
    writer.add_paragraph("")?;

    // Section 3: Different sizes
    writer.add_paragraph("3. Images with Different Sizes")?;
    writer.add_paragraph("Small image:")?;
    writer.add_paragraph("[Image 3 placeholder - small]")?;
    writer.add_paragraph("Large image:")?;
    writer.add_paragraph("[Image 4 placeholder - large]")?;

    writer.add_paragraph("The public DOC writer currently supports the surrounding narrative and layout used to describe inline image positions.")?;
    writer.add_paragraph("[Image 1 placeholder: JPEG, 2x2 inches]")?;
    writer.add_paragraph("[Image 2 placeholder: PNG, 2x2 inches]")?;
    writer.add_paragraph("[Image 3 placeholder: JPEG, 1x1 inch]")?;
    writer.add_paragraph("[Image 4 placeholder: JPEG, 3x3 inches]")?;

    println!("Image-oriented sections configured:");
    println!("  - Image 1: JPEG placeholder, 2x2 inches");
    println!("  - Image 2: PNG placeholder, 2x2 inches");
    println!("  - Image 3: JPEG placeholder, 1x1 inch (small)");
    println!("  - Image 4: JPEG placeholder, 3x3 inches (large)");

    // Save the document
    let output_path = "output/doc_images.doc";
    writer.save(output_path)?;

    println!("\n✅ Document saved to: {}", output_path);
    println!("\n📝 Open this file in Microsoft Word to verify:");
    println!("   - Image placeholder sections should be visible in the document");
    println!("   - Text layout around the image placeholders should match the example structure");
    println!("   - Use this as a baseline for future public image writer support");

    Ok(())
}
