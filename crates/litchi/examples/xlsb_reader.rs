//! Comprehensive XLSB Reader Example
//!
//! Demonstrates XLSB reading through the unified `litchi::open` API: the
//! format is detected from the content, and the workbook view refuses a file
//! that is not a workbook.
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_reader --features xlsb -- xlsb_features.xlsb
//! ```

use std::env;
use std::process;

/// The error followed by its causes. A `litchi::Error` displays its code and
/// a summary; the error of the format crate that failed is its source.
fn with_causes(error: &dyn std::error::Error) -> String {
    let mut text = error.to_string();
    let mut cause = error.source();
    while let Some(error) = cause {
        text.push_str(": ");
        text.push_str(&error.to_string());
        cause = error.source();
    }
    text
}

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <xlsb_file>", args[0]);
        eprintln!("Example: {} xlsb_features.xlsb", args[0]);
        process::exit(1);
    }

    let filename = &args[1];
    println!("{}", "=".repeat(80));
    println!("XLSB Reader - Comprehensive Feature Test");
    println!("{}", "=".repeat(80));
    println!("Reading file: {}\n", filename);

    // Open the file; its format is detected from the content.
    let file = match litchi::open(filename) {
        Ok(file) => file,
        Err(e) => {
            eprintln!("Error opening workbook: {}", with_causes(&e));
            process::exit(1);
        },
    };

    // The workbook view refuses a file of another family and names the
    // format it found.
    if let Err(e) = file.workbook() {
        eprintln!("Error opening workbook: {}", with_causes(&e));
        process::exit(1);
    }

    // Extract all text
    match file.text() {
        Ok(text) => {
            println!("Extracted text from all worksheets:");
            println!("{}", "─".repeat(80));
            println!("{}", text);
        },
        Err(e) => {
            eprintln!("Error extracting text: {}", with_causes(&e));
        },
    }

    println!("\n{}", "=".repeat(80));
    println!("Reading completed successfully!");
    println!("{}", "=".repeat(80));
}
