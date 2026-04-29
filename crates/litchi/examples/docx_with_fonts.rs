//! Example demonstrating font embedding in DOCX files.
//!
//! This example creates a Word document with various fonts and enables
//! font embedding and subsetting. Open the generated file in Microsoft Word
//! to verify that fonts are properly embedded.

use litchi::ooxml::docx::Package;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    println!("Creating DOCX file with embedded fonts...");

    // Create a new document
    let mut pkg = Package::new()?;

    // Enable font embedding with subsetting
    pkg.opc_package_mut().with_font_embedding(true, true);

    // Get mutable document
    let doc = pkg.document_mut()?;

    // Add title
    doc.add_heading("Font Embedding Demo", 1)?;

    // Add paragraphs with different fonts (using fonts available on Linux)
    doc.add_paragraph_with_text("This is the default font.");

    // Add a paragraph with Liberation Serif (Times New Roman alternative)
    let p1 = doc.add_paragraph();
    let run1 = p1.add_run();
    run1.set_text("This text uses Liberation Serif font (similar to Times New Roman).");
    run1.font_name("Liberation Serif");
    run1.font_size(48); // 24pt = 48 half-points

    // Add a paragraph with DejaVu Sans
    let p2 = doc.add_paragraph();
    let run2 = p2.add_run();
    run2.set_text("This text uses DejaVu Sans font with bold styling.");
    run2.font_name("DejaVu Sans");
    run2.font_size(44); // 22pt = 44 half-points
    run2.bold(true);

    // Add a paragraph with DejaVu Sans Mono (monospace)
    let p3 = doc.add_paragraph();
    let run3 = p3.add_run();
    run3.set_text("This text uses DejaVu Sans Mono (monospace font).");
    run3.font_name("DejaVu Sans Mono");
    run3.font_size(40); // 20pt = 40 half-points
    run3.italic(true);

    // Add a heading with Liberation Sans
    doc.add_heading("Section with Liberation Sans", 2)?;

    let p4 = doc.add_paragraph();
    let run4 = p4.add_run();
    run4.set_text("This paragraph demonstrates Liberation Sans with various Unicode characters: ");
    run4.font_name("Liberation Sans");

    let run5 = p4.add_run();
    run5.set_text("αβγδε €£¥ ©®™ ←→↑↓ ≠≤≥");
    run5.font_name("Liberation Sans");
    run5.bold(true);

    // Add information paragraph
    doc.add_paragraph_with_text("\n");
    doc.add_paragraph_with_text("Note: Open this file in Microsoft Word and check File > Info > Inspect Document > Check for Issues > Inspect Document to see embedded fonts.");

    // Save the document
    let output_path = "output_docx_with_fonts.docx";
    pkg.save(output_path)?;

    println!("✓ Successfully created: {}", output_path);
    println!("  Open this file in Microsoft Word to verify font embedding.");
    println!("  The fonts should display correctly even if they're not installed on the system.");

    Ok(())
}
