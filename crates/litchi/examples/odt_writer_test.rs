//! Comprehensive ODT (OpenDocument Text) writing example.
//!
//! This example demonstrates all writing capabilities for ODT files,
//! creating a feature-rich document to showcase the library's capabilities.
//!
//! Run with:
//! ```bash
//! cargo run --example odt_writer_test --features odf --no-default-features
//! ```

#[cfg(feature = "odf")]
use litchi::Result;
#[cfg(feature = "odf")]
use litchi::common::Metadata;
#[cfg(feature = "odf")]
use litchi::odt::Builder;

#[cfg(feature = "odf")]
fn main() -> Result<()> {
    println!("=== ODT Writer Comprehensive Test ===\n");

    let output_file = "odt_writer_test_output.odt";
    println!("📝 Creating comprehensive ODT document: {}", output_file);

    // Create a new document using the ODT builder.
    let mut builder = Builder::new();

    // Set document metadata
    println!("✅ Setting metadata...");
    let metadata = Metadata {
        title: Some("Comprehensive ODT Writer Test Document".to_string()),
        author: Some("Litchi Library Test Suite".to_string()),
        description: Some(
            "This document demonstrates all writing capabilities of the litchi ODT writer module."
                .to_string(),
        ),
        subject: Some("ODT Writer Test".to_string()),
        ..Metadata::default()
    };
    builder.set_metadata(metadata);

    // Add main title heading
    println!("✅ Adding headings and paragraphs...");
    builder.add_heading("ODT Writer Comprehensive Test", 1)?;
    builder.add_paragraph("This document showcases all currently supported ODT writing features in the litchi library.")?;

    // Section 1: Paragraph Features
    builder.add_heading("1. Paragraph Features", 2)?;
    builder
        .add_paragraph("This section demonstrates basic paragraph creation and text content.")?;
    builder.add_paragraph("Multiple consecutive paragraphs can be added easily.")?;
    builder.add_paragraph("Each paragraph maintains its own formatting and style properties.")?;

    // Section 2: Heading Hierarchy
    builder.add_heading("2. Heading Hierarchy", 2)?;
    builder.add_paragraph("The library supports multiple heading levels:")?;
    builder.add_heading("2.1. Level 3 Heading Example", 3)?;
    builder.add_paragraph("This is content under a level 3 heading.")?;
    builder.add_heading("2.2. Another Level 3 Heading", 3)?;
    builder.add_paragraph("More content demonstrating heading hierarchy.")?;
    builder.add_heading("2.2.1. Level 4 Heading Example", 4)?;
    builder.add_paragraph("Even deeper nesting is supported.")?;

    // Section 3: List Features
    builder.add_heading("3. List Features", 2)?;
    builder.add_paragraph("The library supports both ordered and unordered lists:")?;

    println!("✅ Adding lists...");
    // Unordered list
    builder.add_bulleted_list(vec![
        "First unordered item",
        "Second unordered item",
        "Third unordered item with longer text to demonstrate wrapping",
        "Fourth item with special characters: α β γ δ",
    ])?;

    builder.add_paragraph("And ordered lists:")?;

    // Ordered list
    builder.add_numbered_list(vec![
        "First ordered item",
        "Second ordered item",
        "Third ordered item",
        "Fourth ordered item",
        "Fifth ordered item",
    ])?;

    // Section 4: Unicode Support
    builder.add_heading("4. Unicode and Special Characters", 2)?;
    builder.add_paragraph("The library properly handles Unicode and special characters:")?;
    builder.add_paragraph("Mathematical symbols: α β γ δ ε ∫ ∑ ∏ √ ∞ ≈ ≠ ≤ ≥")?;
    builder.add_paragraph("Currency symbols: $ € £ ¥ ₹ ₽ ₩")?;
    builder.add_paragraph("Multilingual text: Hello, 你好, こんにちは, Привет, مرحبا, שלום")?;
    builder.add_paragraph("Emoji support: 😀 🎉 🚀 ⭐ 💯 🔥")?;

    // Section 5: Long Content Test
    builder.add_heading("5. Long Content Test", 2)?;
    builder.add_paragraph(
        "This section tests handling of longer paragraphs with substantial text content. \
        Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor \
        incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud \
        exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute \
        irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla \
        pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia \
        deserunt mollit anim id est laborum.",
    )?;

    builder.add_paragraph(
        "Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque \
        laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi \
        architecto beatae vitae dicta sunt explicabo. Nemo enim ipsam voluptatem quia voluptas \
        sit aspernatur aut odit aut fugit, sed quia consequuntur magni dolores eos qui ratione \
        voluptatem sequi nesciunt.",
    )?;

    // Section 6: Nested Lists and Content
    builder.add_heading("6. Combined Content", 2)?;
    builder.add_paragraph("Combining different elements in sequence:")?;

    builder.add_heading("6.1. Development Workflow", 3)?;
    builder.add_paragraph("A typical software development workflow includes:")?;
    builder.add_numbered_list(vec![
        "Requirements gathering and analysis",
        "Design and architecture planning",
        "Implementation and coding",
        "Testing and quality assurance",
        "Deployment and maintenance",
    ])?;

    // Section 7: Technical Writing
    builder.add_heading("7. Technical Documentation", 2)?;
    builder.add_paragraph("This section demonstrates technical documentation features:")?;

    println!("✅ Adding technical content...");
    builder.add_paragraph("API Methods Available:")?;
    builder.add_bulleted_list(vec![
        "Document::open() - Load documents from file",
        "Document::from_bytes() - Load from memory",
        "text() - Extract plain text",
        "paragraphs() - Get structured paragraphs",
        "tables() - Parse document tables",
    ])?;

    builder.add_paragraph("Programming Language Comparison:")?;
    builder.add_paragraph("Rust offers memory safety, Python provides ease of use, JavaScript enables web development, and Go excels at concurrent programming.")?;

    // Section 8: Multiple Topics
    builder.add_heading("8. Documentation Best Practices", 2)?;
    builder.add_paragraph("When writing technical documentation, consider:")?;

    builder.add_bulleted_list(vec![
        "Clear and concise language",
        "Proper code examples",
        "Visual aids and diagrams",
        "Version information",
        "Troubleshooting guides",
    ])?;

    builder.add_paragraph("Code example best practices:")?;
    builder.add_numbered_list(vec![
        "Include complete, runnable examples",
        "Show both simple and complex use cases",
        "Document expected output",
        "Highlight important patterns",
        "Test all examples before publishing",
    ])?;

    // Section 9: Statistics
    builder.add_heading("9. Document Statistics", 2)?;
    builder.add_paragraph("This comprehensive test document includes:")?;
    builder.add_bulleted_list(vec![
        "Multiple heading levels (1-4)",
        "Various paragraph styles and content",
        "Both ordered and unordered lists",
        "Unicode characters and special symbols",
        "Long-form text content",
        "Nested structures and mixed content",
    ])?;

    // Section 10: Conclusion
    builder.add_heading("10. Conclusion", 2)?;
    builder.add_paragraph(
        "This document successfully demonstrates all currently supported ODT writing features \
        in the litchi library. The document can be opened in LibreOffice Writer, OpenOffice \
        Writer, or any other ODF-compatible word processor.",
    )?;

    builder.add_paragraph(
        "For more information about the litchi library, visit the project repository.",
    )?;

    // Save the document
    println!("💾 Saving document to: {}", output_file);
    builder.save(output_file)?;

    println!("✅ Document saved successfully!");
    println!("\n📊 Document Contents:");
    println!("  - Sections: 10");
    println!("  - Headings: ~20");
    println!("  - Paragraphs: ~30+");
    println!("  - Lists: 9");
    println!("\n=== ODT Writer Test Complete ===");
    println!("✅ Comprehensive ODT file created successfully!");
    println!("📖 Open 'odt_writer_test_output.odt' in LibreOffice Writer to view the result.");

    Ok(())
}

#[cfg(not(feature = "odf"))]
fn main() {
    eprintln!(
        "This example requires the 'odf' feature. Try: cargo run --example odt_writer_test --features odf --no-default-features"
    );
}
