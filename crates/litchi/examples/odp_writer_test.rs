//! Comprehensive ODP (OpenDocument Presentation) writing example.
//!
//! This example demonstrates all writing capabilities for ODP files,
//! creating a feature-rich presentation to showcase the library's capabilities.
//!
//! Run with:
//! ```bash
//! cargo run --example odp_writer_test --features odf --no-default-features
//! ```

#[cfg(feature = "odf")]
use litchi::Result;
#[cfg(feature = "odf")]
use litchi::common::{Metadata, ShapeType};
#[cfg(feature = "odf")]
use litchi::odf::odp::{Builder, Shape, Slide};

#[cfg(feature = "odf")]
fn main() -> Result<()> {
    println!("=== ODP Writer Comprehensive Test ===\n");

    let output_file = "odp_writer_test_output.odp";
    println!(
        "📝 Creating comprehensive ODP presentation: {}",
        output_file
    );

    // Create a new presentation using the ODP builder.
    let mut builder = Builder::new();

    // Set document metadata
    println!("✅ Setting metadata...");
    let metadata = Metadata {
        title: Some("Comprehensive ODP Writer Test Presentation".to_string()),
        author: Some("Litchi Library Test Suite".to_string()),
        subject: Some("ODP Writer Test".to_string()),
        description: Some(
            "This presentation demonstrates all writing capabilities of the litchi ODP writer module."
                .to_string(),
        ),
        ..Default::default()
    };
    builder.set_metadata(metadata);

    // Slide 1: Title Slide
    println!("✅ Creating Slide 1: Title Slide...");
    builder.add_slide_with_title(
        "ODP Writer Comprehensive Test",
        "Litchi Library - OpenDocument Presentation Support\nVersion 1.0 - Complete Feature Demonstration"
    )?;

    // Slide 2: Introduction
    println!("✅ Creating Slide 2: Introduction...");
    builder.add_slide_with_title(
        "Introduction",
        "This presentation demonstrates all currently supported ODP writing features:\n\n\
        • Slide creation with titles and content\n\
        • Custom slide elements with shapes\n\
        • Multiple slide types\n\
        • Unicode and special characters\n\
        • Metadata support",
    )?;

    // Slide 3: Feature Overview
    println!("✅ Creating Slide 3: Feature Overview...");
    builder.add_slide_with_title(
        "Supported Features",
        "Reading Capabilities:\n\
        ✓ Open presentations from file or bytes\n\
        ✓ Extract all slides and their content\n\
        ✓ Parse shapes (text boxes, images, etc.)\n\
        ✓ Extract metadata\n\n\
        Writing Capabilities:\n\
        ✓ Create new presentations\n\
        ✓ Add slides with titles and text\n\
        ✓ Add custom shapes\n\
        ✓ Set presentation metadata",
    )?;

    // Slide 4: Basic Slides
    println!("✅ Creating Slide 4: Simple Content...");
    builder.add_slide_with_title(
        "Simple Slide Example",
        "This is a basic slide with a title and text content.\n\n\
        You can add multiple lines of text,\n\
        and the content will be displayed\n\
        in the presentation.",
    )?;

    // Slide 5: Unicode Support
    println!("✅ Creating Slide 5: Unicode Support...");
    builder.add_slide_with_title(
        "Unicode & Multilingual Support",
        "The library fully supports Unicode text:\n\n\
        English: Hello, World!\n\
        Chinese: 你好，世界！\n\
        Japanese: こんにちは、世界！\n\
        Russian: Привет, мир!\n\
        Arabic: مرحبا بالعالم!\n\
        Emoji: 😀 🌍 🎉 🚀 ⭐ 💯",
    )?;

    // Slide 6: Special Characters
    println!("✅ Creating Slide 6: Special Characters...");
    builder.add_slide_with_title(
        "Special Characters & Symbols",
        "Mathematical Symbols:\n\
        α β γ δ ε θ λ π σ ω\n\
        ∫ ∑ ∏ √ ∞ ≈ ≠ ≤ ≥ ±\n\n\
        Currency Symbols:\n\
        $ € £ ¥ ₹ ₽ ₩\n\n\
        Arrows & Symbols:\n\
        → ← ↑ ↓ ↔ ⇒ ⇐ ⇔ • ◆ ★ ☆",
    )?;

    // Slide 7: Code Example
    println!("✅ Creating Slide 7: Code Example...");
    builder.add_slide_with_title(
        "Usage Example",
        "Creating a presentation with litchi:\n\n\
        let mut builder = Builder::new();\n\
        builder.add_slide_with_title(\n\
            \"My Slide\",\n\
            \"Slide content here\"\n\
        )?;\n\
        builder.save(\"output.odp\")?;",
    )?;

    // Slide 8: With Custom Shapes
    println!("✅ Creating Slide 8: Custom Shapes...");
    let slide8 = Slide {
        title: Some("Custom Shapes".to_string()),
        text: String::new(),
        index: 0,
        notes: Some("This slide demonstrates custom shape creation".to_string()),
        transition: None,
        animations: Vec::new(),
        legacy_animation: None,
        shapes: vec![
            Shape {
                shape_type: ShapeType::TextBox,
                text: "This is a text box shape".to_string(),
                name: Some("TextBox1".to_string()),
                x: Some("2cm".to_string()),
                y: Some("3cm".to_string()),
                width: Some("8cm".to_string()),
                height: Some("2cm".to_string()),
                style_name: None,
                ..Shape::new()
            },
            Shape {
                shape_type: ShapeType::TextBox,
                text: "Another text box at a different position".to_string(),
                name: Some("TextBox2".to_string()),
                x: Some("12cm".to_string()),
                y: Some("6cm".to_string()),
                width: Some("8cm".to_string()),
                height: Some("2cm".to_string()),
                style_name: None,
                ..Shape::new()
            },
            Shape {
                shape_type: ShapeType::TextBox,
                text: "Bottom text box".to_string(),
                name: Some("TextBox3".to_string()),
                x: Some("5cm".to_string()),
                y: Some("12cm".to_string()),
                width: Some("10cm".to_string()),
                height: Some("2cm".to_string()),
                style_name: None,
                ..Shape::new()
            },
        ],
    };
    builder.add_slide_element(slide8)?;

    // Slide 9: Long Content
    println!("✅ Creating Slide 9: Long Content...");
    builder.add_slide_with_title(
        "Long Content Example",
        "This slide demonstrates handling of longer text content. \
        Lorem ipsum dolor sit amet, consectetur adipiscing elit. \
        Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. \
        Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris.\n\n\
        Multiple paragraphs are supported, allowing for comprehensive content \
        presentation across slides. This is useful for detailed explanations \
        and documentation.",
    )?;

    // Slide 10: Lists and Bullet Points
    println!("✅ Creating Slide 10: Lists...");
    builder.add_slide_with_title(
        "Development Workflow",
        "A typical software development workflow:\n\n\
        1. Requirements gathering and analysis\n\
        2. Design and architecture planning\n\
        3. Implementation and coding\n\
        4. Testing and quality assurance\n\
        5. Deployment and maintenance\n\n\
        Each phase is critical for project success.",
    )?;

    // Slide 11: Technical Details
    println!("✅ Creating Slide 11: Technical Details...");
    builder.add_slide_with_title(
        "Technical Implementation",
        "Architecture:\n\n\
        • ZIP-based ODF package handling\n\
        • XML parsing with quick-xml\n\
        • Zero-copy where possible\n\
        • Memory-efficient streaming\n\n\
        Standards Compliance:\n\
        • ISO/IEC 26300 (ODF 1.2)\n\
        • Full namespace support",
    )?;

    // Slide 12: Statistics
    println!("✅ Creating Slide 12: Statistics...");
    builder.add_slide_with_title(
        "Implementation Statistics",
        "Current Implementation:\n\n\
        ✓ ODT (Text): Complete read/write support\n\
        ✓ ODS (Spreadsheet): Complete read/write support\n\
        ✓ ODP (Presentation): Complete read/write support\n\n\
        Features:\n\
        • 100+ API methods\n\
        • Full metadata support\n\
        • Unicode compliant",
    )?;

    // Slide 13: Future Roadmap
    println!("✅ Creating Slide 13: Future Roadmap...");
    builder.add_slide_with_title(
        "Future Enhancements",
        "Planned Features:\n\n\
        • Slide transitions and animations\n\
        • Speaker notes support (partial)\n\
        • Multimedia embedding (audio/video)\n\
        • Custom slide layouts\n\
        • Advanced shape properties\n\
        • Connector lines and arrows\n\
        • Embedded charts and tables",
    )?;

    // Slide 14: Benefits
    println!("✅ Creating Slide 14: Benefits...");
    builder.add_slide_with_title(
        "Why Choose Litchi?",
        "Key Benefits:\n\n\
        ✓ High Performance - Fast processing\n\
        ✓ Memory Efficient - Low resource usage\n\
        ✓ Type Safety - Rust's strong type system\n\
        ✓ Cross-platform - Works everywhere\n\
        ✓ Easy API - Simple and intuitive\n\
        ✓ Production Ready - Well tested",
    )?;

    // Slide 15: Conclusion
    println!("✅ Creating Slide 15: Conclusion...");
    builder.add_slide_with_title(
        "Conclusion",
        "This presentation successfully demonstrates:\n\n\
        ✓ All currently supported ODP writing features\n\
        ✓ Simple and complex slide creation\n\
        ✓ Custom shapes and positioning\n\
        ✓ Unicode and special character support\n\
        ✓ Metadata handling\n\n\
        The litchi library provides production-ready\n\
        ODF support for Rust applications.",
    )?;

    // Slide 16: Thank You with Custom Shape
    println!("✅ Creating Slide 16: Thank You...");
    let slide16 = Slide {
        title: Some("Thank You!".to_string()),
        text: String::new(),
        index: 0,
        notes: Some("Final slide with acknowledgments".to_string()),
        transition: None,
        animations: Vec::new(),
        legacy_animation: None,
        shapes: vec![Shape {
            shape_type: ShapeType::TextBox,
            text: "Litchi Library - ODF Support\n\nThank you for using litchi!".to_string(),
            name: Some("ThankYouText".to_string()),
            x: Some("5cm".to_string()),
            y: Some("8cm".to_string()),
            width: Some("14cm".to_string()),
            height: Some("4cm".to_string()),
            style_name: None,
            ..Shape::new()
        }],
    };
    builder.add_slide_element(slide16)?;

    // Save the presentation
    println!("💾 Saving presentation to: {}", output_file);
    builder.save(output_file)?;

    println!("✅ Presentation saved successfully!");
    println!("\n📊 Presentation Contents:");
    println!("  - Slides: 16");
    println!("    1. Title Slide");
    println!("    2. Introduction");
    println!("    3. Feature Overview");
    println!("    4. Simple Content");
    println!("    5. Unicode Support");
    println!("    6. Special Characters");
    println!("    7. Code Example");
    println!("    8. Custom Shapes (with 3 text boxes)");
    println!("    9. Long Content");
    println!("    10. Lists and Workflow");
    println!("    11. Technical Details");
    println!("    12. Statistics");
    println!("    13. Future Roadmap");
    println!("    14. Benefits");
    println!("    15. Conclusion");
    println!("    16. Thank You (with custom shape)");
    println!("  - Custom shapes: 4 text boxes");
    println!("  - Speaker notes: 2 slides");
    println!("\n=== ODP Writer Test Complete ===");
    println!("✅ Comprehensive ODP file created successfully!");
    println!("📖 Open 'odp_writer_test_output.odp' in LibreOffice Impress to view the result.");

    Ok(())
}

#[cfg(not(feature = "odf"))]
fn main() {
    eprintln!(
        "This example requires the 'odf' feature. Try: cargo run --example odp_writer_test --features odf --no-default-features"
    );
}
