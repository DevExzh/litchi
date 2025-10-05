/// Example: Parse a legacy Word document (.doc)
///
/// This example demonstrates how to use the Litchi library to read
/// and extract text from a .doc file.
///
/// Usage:
///   cargo run --example parse_doc -- <path_to_doc_file>
use litchi::ole::doc::Package;
use std::env;
use std::process;

fn main() {
    let args: Vec<String> = env::args().collect();

    if args.len() < 2 {
        eprintln!("Usage: {} [path_to_doc_file]", args[0]);
        eprintln!("\nExamples:");
        eprintln!("  cargo run --example parse_doc -- document.doc");
        eprintln!("  cargo run --example parse_doc    # uses test.doc");
        process::exit(1);
    }

    let file_path = if args.len() > 1 {
        &args[1]
    } else {
        "test.doc"
    };

    println!("Opening DOC file: {}", file_path);
    println!("{}", "=".repeat(60));

    // Open the DOC file
    let mut package = match Package::open(file_path) {
        Ok(pkg) => pkg,
        Err(e) => {
            eprintln!("Error opening file: {}", e);
            process::exit(1);
        }
    };

    // Get the main document
    let document = match package.document() {
        Ok(doc) => doc,
        Err(e) => {
            eprintln!("Error reading document: {}", e);
            process::exit(1);
        }
    };

    // Display document information
    println!("\n📄 Document Information:");
    println!("{}", "-".repeat(60));

    // Get FIB information
    let fib = document.fib();
    println!("  Format Version:  0x{:04X}", fib.version());
    println!("  Table Stream:    {}", if fib.which_table_stream() { "1Table" } else { "0Table" });
    println!("  Encrypted:       {}", if fib.is_encrypted() { "Yes ⚠️" } else { "No" });
    println!("  Language ID:     0x{:04X}", fib.language_id());

    // Extract and display text
    println!("\n📝 Document Text:");
    println!("{}", "-".repeat(60));

    match document.text() {
        Ok(text) => {
            if text.is_empty() {
                println!("  (Document is empty or text extraction not supported)");
            } else {
                let lines: Vec<&str> = text.lines().collect();
                println!("  Total characters: {}", text.len());
                println!("  Total lines:      {}", lines.len());
                println!("\n  Content preview (first 10 lines):");
                println!("  {}", "-".repeat(58));
                for (i, line) in lines.iter().take(10).enumerate() {
                    if line.trim().is_empty() {
                        println!("  {:3}: (empty line)", i + 1);
                    } else {
                        let preview = if line.len() > 70 {
                            format!("{}...", &line[..67])
                        } else {
                            line.to_string()
                        };
                        println!("  {:3}: {}", i + 1, preview);
                    }
                }

                if lines.len() > 10 {
                    println!("  ... ({} more lines)", lines.len() - 10);
                }
            }
        }
        Err(e) => {
            eprintln!("  Error extracting text: {}", e);
        }
    }

    // Display paragraph count
    println!("\n📋 Document Structure:");
    println!("{}", "-".repeat(60));

    match document.paragraph_count() {
        Ok(count) => {
            println!("  Paragraphs:  {}", count);
        }
        Err(e) => {
            eprintln!("  Error counting paragraphs: {}", e);
        }
    }

    match document.table_count() {
        Ok(count) => {
            println!("  Tables:      {}", count);
        }
        Err(e) => {
            eprintln!("  Error counting tables: {}", e);
        }
    }

    // Access individual paragraphs
    println!("\n📑 Paragraphs:");
    println!("{}", "-".repeat(60));

    match document.paragraphs() {
        Ok(paragraphs) => {
            let display_count = paragraphs.len().min(5);
            for (i, para) in paragraphs.iter().take(display_count).enumerate() {
                match para.text() {
                    Ok(text) => {
                        if !text.trim().is_empty() {
                            let preview = if text.len() > 60 {
                                format!("{}...", &text[..57])
                            } else {
                                text.to_string()
                            };
                            println!("  [{}] {}", i + 1, preview);
                        }
                    }
                    Err(e) => {
                        eprintln!("  [{}] Error: {}", i + 1, e);
                    }
                }
            }

            if paragraphs.len() > display_count {
                println!("  ... ({} more paragraphs)", paragraphs.len() - display_count);
            }
        }
        Err(e) => {
            eprintln!("  Error reading paragraphs: {}", e);
        }
    }

    println!("\n{}", "=".repeat(60));
    println!("✅ Done!");
}

