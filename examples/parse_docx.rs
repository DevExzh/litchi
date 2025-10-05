/// Example demonstrating how to parse and extract text content from a .docx file
/// using the litchi OOXML DOCX parser.
///
/// This example shows:
/// - Opening a DOCX package (Office Open XML)
/// - Extracting document text content
/// - Accessing paragraphs and tables
/// - Displaying document structure and metadata
use litchi::ooxml::docx::Package;
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Get the file path from command line arguments
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} [path-to-docx-file]", args[0]);
        eprintln!("\nExamples:");
        eprintln!("  cargo run --example parse_docx -- document.docx");
        eprintln!("  cargo run --example parse_docx    # uses test.docx");
        std::process::exit(1);
    }

    let file_path = if args.len() > 1 {
        &args[1]
    } else {
        "test.docx"
    };

    println!("Opening DOCX document: {}", file_path);
    println!("{}", "=".repeat(60));

    // Open the DOCX package
    let package = Package::open(file_path)?;
    let doc = package.document()?;

    // Extract and display document text content
    println!("\n📝 Document Text Content:");
    println!("{}", "-".repeat(60));

    match doc.text() {
        Ok(text) => {
            if text.is_empty() {
                println!("  (Document is empty or no text content found)");
            } else {
                let lines: Vec<&str> = text.lines().collect();
                println!("  Total characters: {}", text.len());
                println!("  Total lines:      {}", lines.len());
                println!("\n  Content preview (first 20 lines):");
                println!("  {}", "-".repeat(58));
                for (i, line) in lines.iter().take(20).enumerate() {
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

                if lines.len() > 20 {
                    println!("  ... ({} more lines)", lines.len() - 20);
                }
            }
        }
        Err(e) => {
            eprintln!("  Error extracting text: {}", e);
        }
    }

    // Display document structure
    println!("\n📋 Document Structure:");
    println!("{}", "-".repeat(60));

    match doc.paragraph_count() {
        Ok(count) => {
            println!("  Paragraphs:  {}", count);
        }
        Err(e) => {
            eprintln!("  Error counting paragraphs: {}", e);
        }
    }

    match doc.table_count() {
        Ok(count) => {
            println!("  Tables:      {}", count);
        }
        Err(e) => {
            eprintln!("  Error counting tables: {}", e);
        }
    }

    // Access and display paragraphs
    println!("\n📑 Paragraphs:");
    println!("{}", "-".repeat(60));

    match doc.paragraphs() {
        Ok(paragraphs) => {
            let display_count = paragraphs.len().min(10);
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

    // Access and display tables
    println!("\n📊 Tables:");
    println!("{}", "-".repeat(60));

    match doc.tables() {
        Ok(tables) => {
            for (i, table) in tables.iter().enumerate() {
                match table.row_count() {
                    Ok(row_count) => {
                        println!("  Table {}: {} rows", i + 1, row_count);
                        // Show first few cells of first row
                        if row_count > 0 {
                            match table.rows() {
                                Ok(rows) => {
                                    if let Some(first_row) = rows.first() {
                                        match first_row.cells() {
                                            Ok(cells) => {
                                                let cell_count = cells.len().min(3);
                                                print!("    First row cells: ");
                                                for (j, cell) in cells.iter().take(cell_count).enumerate() {
                                                    match cell.text() {
                                                        Ok(text) => {
                                                            let preview = if text.len() > 20 {
                                                                format!("\"{}...\"", &text[..17])
                                                            } else {
                                                                format!("\"{}\"", text)
                                                            };
                                                            if j > 0 { print!(", "); }
                                                            print!("{}", preview);
                                                        }
                                                        Err(_) => print!("(error)")
                                                    }
                                                }
                                                println!();
                                                if cells.len() > 3 {
                                                    println!("    ... ({} more cells)", cells.len() - 3);
                                                }
                                            }
                                            Err(_) => {
                                                println!("    (Error accessing cells)");
                                            }
                                        }
                                    }
                                }
                                Err(_) => {
                                    println!("    (Error accessing rows)");
                                }
                            }
                        }
                    }
                    Err(_) => {
                        println!("  Table {}: (Error counting rows)", i + 1);
                    }
                }
            }
        }
        Err(e) => {
            eprintln!("  Error reading tables: {}", e);
        }
    }

    // Show package information
    println!("\n📦 Package Information:");
    println!("{}", "-".repeat(60));
    let opc = package.opc_package();
    println!("  Total parts: {}", opc.part_count());

    // Performance statistics
    println!("\n⚡ Performance Notes:");
    println!("   - Uses memchr for fast string searching");
    println!("   - Uses atoi_simd for fast integer parsing");
    println!("   - Uses quick-xml for zero-copy XML parsing");
    println!("   - Minimal allocations through borrowing");

    Ok(())
}
