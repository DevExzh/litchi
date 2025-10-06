/// Example demonstrating the unified API for Word documents and PowerPoint presentations.
///
/// This example shows how to use the new high-level API that automatically detects
/// file formats and provides a consistent interface for both legacy (.doc, .ppt)
/// and modern (.docx, .pptx) formats.

use litchi::{Document, Presentation};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== Litchi Unified API Example ===\n");

    // Example 1: Word Document
    println!("1. Reading Word Document");
    println!("   - Format is automatically detected");
    println!("   - Works with both .doc and .docx files\n");

    // Open a document - format is auto-detected
    let doc = Document::open("test.doc")?;

    // Extract all text
    let text = doc.text()?;
    println!("Document text preview:");
    println!("{}\n", text.chars().take(200).collect::<String>());

    // Get paragraph count
    let para_count = doc.paragraph_count()?;
    println!("Total paragraphs: {}\n", para_count);

    // Access individual paragraphs
    println!("First 3 paragraphs:");
    for (i, para) in doc.paragraphs()?.iter().take(3).enumerate() {
        println!("  Paragraph {}: {}", i + 1, para.text()?);
    }
    println!();

    // Example 2: PowerPoint Presentation
    println!("2. Reading PowerPoint Presentation");
    println!("   - Format is automatically detected");
    println!("   - Works with both .ppt and .pptx files\n");

    // Open a presentation - format is auto-detected
    let pres = Presentation::open("test.ppt")?;

    // Get slide count
    let slide_count = pres.slide_count()?;
    println!("Total slides: {}\n", slide_count);

    // Access individual slides
    println!("Slide contents:");
    for (i, slide) in pres.slides()?.iter().enumerate() {
        println!("  Slide {}: {}", i + 1, slide.text()?);
        
        // Get slide name (if available - only for .pptx)
        if let Ok(Some(name)) = slide.name() {
            println!("    Name: {}", name);
        }
    }
    println!();

    // Extract all presentation text
    let pres_text = pres.text()?;
    println!("Full presentation text preview:");
    println!("{}\n", pres_text.chars().take(300).collect::<String>());

    println!("=== Example Complete ===");
    Ok(())
}

