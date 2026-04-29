//! Example demonstrating hyperlinks in DOC files
//!
//! This example creates a DOC file with clickable hyperlinks using
//! Word field codes (HYPERLINK). Each link is rendered as blue underlined
//! text that opens in a browser when clicked.
//!
//! Run with: cargo run --example doc_writer_hyperlinks

use litchi::ole::doc::writer::{DocWriter, ParagraphFormatting};
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    println!("Creating DOC file with hyperlinks...");

    let mut writer = DocWriter::new();

    // Title
    writer.add_paragraph("Hyperlink Examples")?;
    writer.add_paragraph("")?;

    // URL hyperlinks using DocWriter::add_hyperlink (creates proper field codes)
    writer.add_paragraph("1. Web Links")?;
    writer.add_hyperlink(
        "Visit the Rust website",
        "https://www.rust-lang.org",
        ParagraphFormatting::default(),
    )?;
    writer.add_hyperlink(
        "Read the Rust Book",
        "https://doc.rust-lang.org/book/",
        ParagraphFormatting::default(),
    )?;
    writer.add_hyperlink(
        "Browse crates on crates.io",
        "https://crates.io",
        ParagraphFormatting::default(),
    )?;
    writer.add_paragraph("")?;

    // Email hyperlinks
    writer.add_paragraph("2. Email Links")?;
    writer.add_hyperlink(
        "Send us an email",
        "mailto:hello@example.com",
        ParagraphFormatting::default(),
    )?;
    writer.add_hyperlink(
        "Contact support",
        "mailto:support@example.com",
        ParagraphFormatting::default(),
    )?;
    writer.add_paragraph("")?;

    // More body text
    writer.add_paragraph("3. Mixed Content")?;
    writer.add_paragraph("This paragraph has no hyperlink, just regular text.")?;
    writer.add_hyperlink(
        "GitHub - where the world builds software",
        "https://github.com",
        ParagraphFormatting::default(),
    )?;
    writer.add_paragraph("Another regular paragraph after the link.")?;

    println!("Hyperlinks configured:");
    println!("  - 3 URL links");
    println!("  - 2 Email links");
    println!("  - 1 Additional URL link");

    // Save the document
    let output_path = "output/doc_hyperlinks.doc";
    writer.save(output_path)?;

    println!("\n Document saved to: {}", output_path);
    println!("\n Open this file in Microsoft Word to verify:");
    println!("   - Hyperlinks should be underlined and blue");
    println!("   - Ctrl+Click on links to open in browser/email client");
    println!("   - Right-click > Edit Hyperlink to see destinations");

    Ok(())
}
