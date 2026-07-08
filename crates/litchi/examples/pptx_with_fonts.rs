//! Example demonstrating font embedding in PPTX files.
//!
//! This example creates a PowerPoint presentation with various fonts and enables
//! font embedding and subsetting. Open the generated file in Microsoft PowerPoint
//! to verify that fonts are properly embedded.

use litchi::ooxml::pptx::Package;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    println!("Creating PPTX file with embedded fonts...");

    // Create a new presentation
    let mut pkg = Package::new()?;

    // Enable font embedding with subsetting
    pkg.opc_package_mut().with_font_embedding(true, true);

    // Get mutable presentation
    let pres = pkg.presentation_mut()?;

    // Slide 1: Title slide
    let slide1 = pres.add_slide()?;
    slide1.set_title("Font Embedding Demo");
    slide1.set_notes(
        "This presentation demonstrates font embedding in PowerPoint files using Linux fonts.",
    );

    // Slide 2: Content slide with Liberation Serif
    let slide2 = pres.add_slide()?;
    slide2.set_title("Liberation Serif Font");
    slide2.add_text_box(
        "This text box uses Liberation Serif font (similar to Times New Roman). The font should be embedded and display correctly even on systems without this font installed.",
        1000000,  // x: 1 inch (914400 EMU = 1 inch)
        2000000,  // y: 2 inches
        8000000,  // width: 8 inches
        1500000,  // height: 1.5 inches
    )
    .font("Liberation Serif")
    .font_size(18.0);

    // Slide 3: Content slide with DejaVu Sans
    let slide3 = pres.add_slide()?;
    slide3.set_title("DejaVu Sans Font");
    slide3.add_text_box(
        "This text box demonstrates DejaVu Sans font with styling. DejaVu Sans is a modern sans-serif font commonly available on Linux systems.",
        1000000,
        2000000,
        8000000,
        1500000,
    )
    .font("DejaVu Sans")
    .font_size(18.0)
    .bold(true);

    // Slide 4: Content slide with DejaVu Sans Mono
    let slide4 = pres.add_slide()?;
    slide4.set_title("DejaVu Sans Mono (Monospace)");
    slide4.add_text_box(
        "This is DejaVu Sans Mono - a monospace font:\n  Code example: fn main() {\n    println!(\"Hello, World!\");\n  }",
        1000000,
        2000000,
        8000000,
        2000000,
    )
    .font("DejaVu Sans Mono")
    .font_size(16.0);

    // Slide 5: Content slide with Liberation Sans and Unicode characters
    let slide5 = pres.add_slide()?;
    slide5.set_title("Liberation Sans with Unicode");
    slide5.add_text_box(
        "Liberation Sans font with special characters:\nGreek: αβγδε\nCurrency: €£¥\nSymbols: ©®™\nArrows: ←→↑↓\nMath: ≠≤≥",
        1000000,
        2000000,
        8000000,
        2500000,
    )
    .font("Liberation Sans")
    .font_size(20.0);

    // Slide 6: Summary slide
    let slide6 = pres.add_slide()?;
    slide6.set_title("Verification Instructions");
    slide6.add_text_box(
        "To verify font embedding:\n\n1. Open this file in Microsoft PowerPoint\n2. Go to File > Info > Properties\n3. Check embedded fonts information\n4. The fonts should display correctly even if not installed",
        1000000,
        2000000,
        8000000,
        3000000,
    );

    // Save the presentation
    let output_path = "output_pptx_with_fonts.pptx";
    pkg.save(output_path)?;

    println!("✓ Successfully created: {}", output_path);
    println!("  Open this file in Microsoft PowerPoint to verify font embedding.");
    println!("  The fonts should display correctly even if they're not installed on the system.");

    Ok(())
}
