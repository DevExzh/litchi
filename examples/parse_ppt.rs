/// Example: Parse a PowerPoint (.ppt) presentation
///
/// This example demonstrates how to use the Litchi library to parse
/// legacy PowerPoint presentations in the binary .ppt format.

use litchi::ole::ppt::Package;
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();

    if args.len() < 2 {
        eprintln!("Usage: {} [path_to_ppt_file]", args[0]);
        eprintln!("\nExamples:");
        eprintln!("  cargo run --example parse_ppt -- presentation.ppt");
        eprintln!("  cargo run --example parse_ppt    # uses test.ppt");
        std::process::exit(1);
    }

    // Open a .ppt file (use command line arg or default to test.ppt)
    let file_path = if args.len() > 1 {
        &args[1]
    } else {
        "test.ppt"
    };

    let mut pkg = Package::open(file_path)?;
    println!("Successfully opened PPT file");

    // Get the main presentation
    let pres = pkg.presentation()?;
    println!("Successfully loaded presentation");

    // Extract all text (placeholder implementation for now)
    let text = pres.text()?;
    println!("Presentation text: {}", text);

    // Get slide count (placeholder implementation for now)
    let slide_count = pres.slide_count()?;
    println!("Number of slides: {}", slide_count);

    // Access slides (placeholder implementation for now)
    for slide in pres.slides()? {
        println!("Slide text: {}", slide.text()?);
        println!("Slide has {} shapes", slide.shape_count()?);

        // Demonstrate placeholder access methods
        println!("Slide placeholder methods available:");
        println!("- get_placeholder(idx)");
        println!("- get_placeholders_by_type(type)");
        println!("- placeholders()");

        // Show available placeholder types
        println!("Available placeholder types:");
        println!("- Title, Body, CenterTitle, SubTitle");
        println!("- Chart, Table, Picture, Object");
        println!("- Header, Footer, SlideNumber, DateAndTime");
    }

    // Get metadata if available
    if let Some(metadata) = pres.metadata()? {
        if let Some(title) = &metadata.title {
            println!("Title: {}", title);
        }
        if let Some(author) = &metadata.author {
            println!("Author: {}", author);
        }
    }

    println!("\nNote: Full shape and placeholder parsing is implemented");
    println!("but requires actual PPT binary format decoding for complete functionality.");

    Ok(())
}
