/// Example: Parse a .docx file using the OOXML API
///
/// This example demonstrates how to use the high-level OOXML API to open
/// and extract information from a Word document.
use litchi::ooxml::docx::Package;
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Get the file path from command line arguments
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <path-to-docx-file>", args[0]);
        eprintln!("Example: {} document.docx", args[0]);
        std::process::exit(1);
    }

    let file_path = &args[1];
    println!("Opening Word document: {}", file_path);
    println!("{}", "=".repeat(60));

    // Open the .docx package
    let package = Package::open(file_path)?;
    println!("✓ Package opened successfully");

    // Get the main document
    let document = package.document()?;
    println!("✓ Document loaded");
    println!();

    // Extract document statistics
    println!("Document Statistics:");
    println!("{}", "-".repeat(60));

    let para_count = document.paragraph_count()?;
    println!("  Paragraphs: {}", para_count);

    let table_count = document.table_count()?;
    println!("  Tables:     {}", table_count);
    println!();

    // Extract all text content
    println!("Text Content:");
    println!("{}", "-".repeat(60));
    let text = document.text()?;

    if text.is_empty() {
        println!("  (Document contains no text)");
    } else {
        // Show first 500 characters
        let preview = if text.len() > 500 {
            format!("{}...", &text[..500])
        } else {
            text.clone()
        };
        println!("{}", preview);

        if text.len() > 500 {
            println!();
            println!("  Total characters: {}", text.len());
            println!("  (showing first 500 characters)");
        }
    }

    println!();
    println!("{}", "=".repeat(60));
    println!("✓ Done");

    Ok(())
}
