//! Comprehensive XLSB Reader Example
//!
//! Demonstrates XLSB reading using the unified Workbook API.
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_reader --features ooxml --no-default-features xlsb_features.xlsb
//! ```

use litchi::sheet::Workbook;
use std::env;
use std::process;

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

    // Open the workbook
    let workbook = match Workbook::open(filename) {
        Ok(wb) => wb,
        Err(e) => {
            eprintln!("Error opening workbook: {}", e);
            process::exit(1);
        },
    };

    // Extract all text
    match workbook.text() {
        Ok(text) => {
            println!("Extracted text from all worksheets:");
            println!("{}", "─".repeat(80));
            println!("{}", text);
        },
        Err(e) => {
            eprintln!("Error extracting text: {}", e);
        },
    }

    println!("\n{}", "=".repeat(80));
    println!("Reading completed successfully!");
    println!("{}", "=".repeat(80));
}
