#![allow(
    clippy::print_stdout,
    reason = "this command-line example intentionally prints its results"
)]

//! Example demonstrating PPT file writing with the Litchi library
//!
//! NOTE: This example demonstrates the API but will not work until
//! the PPT writer implementation is complete. See `OLE_WRITE_SUPPORT_STATUS.md`
//! for implementation status.
//!
//! Run with: `cargo run --example ppt_write_example`
use litchi_ppt::Writer;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating a new PPT file...");

    // Create a new widescreen presentation (16:9)
    let mut writer = Writer::new_widescreen();

    // Set presentation properties
    writer.set_property("Title", "Sample Presentation");
    writer.set_property("Author", "Litchi Example");
    writer.set_property("Subject", "Demonstrating PPT writing");

    // Add title slide
    let slide1 = writer.add_slide()?;
    writer.add_textbox(slide1, 50, 50, 600, 100, "Welcome to Litchi")?;
    writer.add_textbox(
        slide1,
        50,
        200,
        600,
        50,
        "High-Performance Office File Parsing",
    )?;

    // Add content slide
    let slide2 = writer.add_slide()?;
    writer.add_textbox(slide2, 50, 30, 600, 50, "Features")?;
    writer.add_textbox(
        slide2,
        50,
        100,
        600,
        200,
        "• Fast, safe, and idiomatic Rust\n\
         • Support for legacy formats (DOC, XLS, PPT)\n\
         • Support for modern formats (DOCX, XLSX, PPTX)\n\
         • Zero-copy parsing where possible\n\
         • Production-ready quality",
    )?;

    // Add slide with shapes
    let slide3 = writer.add_slide()?;
    writer.add_textbox(slide3, 50, 30, 600, 50, "Architecture")?;
    writer.add_rectangle(slide3, 100, 100, 200, 150)?;
    writer.add_rectangle(slide3, 350, 100, 200, 150)?;
    writer.add_textbox(slide3, 100, 120, 200, 50, "Parser Layer")?;
    writer.add_textbox(slide3, 350, 120, 200, 50, "Writer Layer")?;

    // Add conclusion slide
    let slide4 = writer.add_slide()?;
    writer.add_textbox(slide4, 50, 100, 600, 200, "Thank You!")?;
    writer.set_slide_notes(
        slide4,
        "For more information, visit the project repository and documentation.",
    )?;

    println!("Created {} slides", writer.slide_count());

    // Demonstrate slide reordering
    println!("Reordering slides...");
    writer.move_slide(3, 1)?; // Move conclusion to second position

    // Save the file
    println!("Saving to output.ppt...");
    writer.save("output.ppt")?;

    println!("✅ PPT file created successfully!");
    println!("   - 4 slides with various content");
    println!("   - Text boxes and shapes");
    println!("   - Slide notes");
    println!("   - Slide reordering demonstrated");

    Ok(())
}
