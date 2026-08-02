//! Example demonstrating DOC file writing with the Litchi library
//!
//! NOTE: This example demonstrates the API but will not work until
//! the DOC writer implementation is complete. See OLE_WRITE_SUPPORT_STATUS.md
//! for implementation status.
//!
//! Run with: cargo run --example doc_write_example
use litchi::doc::{CharacterFormatting, DocWriter, ParagraphFormatting};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating a new DOC file...");

    // Create a new DOC writer
    let mut writer = DocWriter::new();

    // Set document properties
    writer.set_property("Title", "Sample Document");
    writer.set_property("Author", "Litchi Example");
    writer.set_property("Subject", "Demonstrating DOC writing");

    // Add a title paragraph with formatting
    let title_char = CharacterFormatting {
        bold: Some(true),
        font_size: Some(28), // 14pt (font size in half-points)
        ..Default::default()
    };
    let title_para = ParagraphFormatting {
        alignment: Some(1), // Center
        ..Default::default()
    };

    writer.add_paragraph_with_format("Project Report", title_char, title_para)?;

    // Add a normal paragraph
    writer.add_paragraph(
        "This document demonstrates the DOC writing capabilities of the Litchi library.",
    )?;

    // Add a section header
    let header_char = CharacterFormatting {
        bold: Some(true),
        font_size: Some(24), // 12pt
        ..Default::default()
    };

    writer.add_paragraph_with_format(
        "Introduction",
        header_char,
        ParagraphFormatting::default(),
    )?;

    // Add body paragraphs
    writer.add_paragraph(
        "The Litchi library provides high-performance parsing and writing of Office file formats. \
         This includes legacy formats like DOC, XLS, and PPT.",
    )?;

    writer.add_paragraph(
        "The implementation follows Microsoft's official specifications and is designed for \
         production use with emphasis on performance, safety, and correctness.",
    )?;

    // Add a table
    let table_idx = writer.add_table(3, 2)?;

    // Set table headers
    writer.set_table_cell_text(table_idx, 0, 0, "Feature")?;
    writer.set_table_cell_text(table_idx, 0, 1, "Status")?;

    // Set table data
    writer.set_table_cell_text(table_idx, 1, 0, "Text Extraction")?;
    writer.set_table_cell_text(table_idx, 1, 1, "Complete")?;

    writer.set_table_cell_text(table_idx, 2, 0, "File Writing")?;
    writer.set_table_cell_text(table_idx, 2, 1, "In Progress")?;

    // Add conclusion
    let conclusion_char = CharacterFormatting {
        italic: Some(true),
        ..Default::default()
    };

    writer.add_paragraph_with_format(
        "For more information, visit the project documentation.",
        conclusion_char,
        ParagraphFormatting::default(),
    )?;

    // Save the file
    println!("Saving to output.doc...");
    writer.save("output.doc")?;

    println!("✅ DOC file created successfully!");
    println!("   - Multiple paragraphs with formatting");
    println!("   - Table with data");
    println!("   - Document properties");

    Ok(())
}
