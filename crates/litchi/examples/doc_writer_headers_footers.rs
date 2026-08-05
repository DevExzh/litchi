//! Example demonstrating headers and footers in DOC files
//!
//! This example creates a DOC file with:
//! - Different headers for odd/even pages
//! - Different footers for odd/even pages
//! - First page header/footer
//!
//! Run with: cargo run --example doc_writer_headers_footers

use litchi_doc::writer::DocWriter;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    println!("Creating DOC file with headers and footers...");

    // Create a new DOC writer
    let mut writer = DocWriter::new();

    // Add some content to the document (need enough for multiple pages)
    writer.add_paragraph("Title Page")?;
    writer.add_paragraph("This is the first page with a special header and footer.")?;
    writer.add_paragraph("")?;
    for _ in 0..10 {
        writer.add_paragraph("Line to fill the first page...")?;
    }
    writer.add_paragraph("")?;

    writer.add_paragraph("Page 2 - Odd Page")?;
    writer.add_paragraph("This page should show the odd page header and footer.")?;
    writer.add_paragraph("")?;
    for _ in 0..10 {
        writer.add_paragraph("Line to fill the second page...")?;
    }
    writer.add_paragraph("")?;

    writer.add_paragraph("Page 3 - Even Page")?;
    writer.add_paragraph("This page should show the even page header and footer.")?;
    writer.add_paragraph("")?;
    for _ in 0..10 {
        writer.add_paragraph("Line to fill the third page...")?;
    }
    writer.add_paragraph("")?;

    writer.add_paragraph("Page 4 - Odd Page Again")?;
    writer.add_paragraph("This page should show the odd page header and footer again.")?;
    for _ in 0..5 {
        writer.add_paragraph("More content on the fourth page...")?;
    }

    // Set headers and footers using DocWriter API
    writer.set_first_header("=== FIRST PAGE HEADER ===");
    writer.set_odd_header("Document Title - Odd Pages");
    writer.set_even_header("Document Title - Even Pages");

    writer.set_first_footer("=== Title Page ===");
    writer.set_odd_footer("Page [odd] - Company Name");
    writer.set_even_footer("Company Name - Page [even]");

    println!("Headers and footers configured:");
    println!("  ✓ First page header: '=== FIRST PAGE HEADER ==='");
    println!("  ✓ Odd page header: 'Document Title - Odd Pages'");
    println!("  ✓ Even page header: 'Document Title - Even Pages'");
    println!("  ✓ First page footer: '=== Title Page ==='");
    println!("  ✓ Odd page footer: 'Page [odd] - Company Name'");
    println!("  ✓ Even page footer: 'Company Name - Page [even]'");

    // Save the document
    let output_path = "output/doc_headers_footers.doc";
    writer.save(output_path)?;

    println!("\n✅ Document saved to: {}", output_path);
    println!("\n📝 Open this file in Microsoft Word to verify:");
    println!("   - View > Header and Footer to see different headers/footers");
    println!("   - Navigate through pages to see odd/even variations");
    println!("   - First page should have special header/footer");

    Ok(())
}
