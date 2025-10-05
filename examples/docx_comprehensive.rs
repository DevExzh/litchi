/// Comprehensive example of docx parsing with full paragraph and table iteration.
///
/// This example demonstrates:
/// - Paragraph iteration with text extraction
/// - Run iteration with formatting properties (bold, italic, underline)
/// - Table iteration with cell access
/// - Font properties (name, size)
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
    println!("{}", "=".repeat(80));

    // Open the .docx package
    let package = Package::open(file_path)?;
    let document = package.document()?;

    println!("✓ Document loaded successfully\n");

    // ===== DOCUMENT STATISTICS =====
    println!("Document Statistics:");
    println!("{}", "-".repeat(80));

    let para_count = document.paragraph_count()?;
    let table_count = document.table_count()?;
    println!("  Total Paragraphs: {}", para_count);
    println!("  Total Tables:     {}", table_count);
    println!();

    // ===== PARAGRAPH ITERATION =====
    println!("Paragraphs with Formatting:");
    println!("{}", "-".repeat(80));

    let paragraphs = document.paragraphs()?;
    for (para_idx, para) in paragraphs.iter().enumerate().take(5) {
        let text = para.text()?;
        if text.trim().is_empty() {
            continue;
        }

        println!(
            "\nParagraph #{}: {}",
            para_idx + 1,
            if text.len() > 60 {
                format!("{}...", &text[..60])
            } else {
                text.clone()
            }
        );

        // Iterate through runs
        let runs = para.runs()?;
        if runs.len() > 0 {
            println!("  Runs: {}", runs.len());
            for (run_idx, run) in runs.iter().enumerate().take(3) {
                let run_text = run.text()?;
                if run_text.trim().is_empty() {
                    continue;
                }

                // Get formatting properties
                let bold = run.bold()?;
                let italic = run.italic()?;
                let underline = run.underline()?;
                let font_name = run.font_name()?;
                let font_size = run.font_size()?;

                // Build formatting description
                let mut format_parts = Vec::new();
                if let Some(true) = bold {
                    format_parts.push("bold".to_string());
                }
                if let Some(true) = italic {
                    format_parts.push("italic".to_string());
                }
                if let Some(true) = underline {
                    format_parts.push("underline".to_string());
                }
                if let Some(name) = &font_name {
                    format_parts.push(format!("font: {}", name));
                }
                if let Some(size) = font_size {
                    format_parts.push(format!("size: {}pt", size as f32 / 2.0));
                }

                let format_str = if format_parts.is_empty() {
                    "default".to_string()
                } else {
                    format_parts.join(", ")
                };

                let display_text = if run_text.len() > 30 {
                    format!("{}...", &run_text[..30])
                } else {
                    run_text
                };

                println!(
                    "    Run #{}: \"{}\" [{}]",
                    run_idx + 1,
                    display_text,
                    format_str
                );
            }
            if runs.len() > 3 {
                println!("    ... and {} more runs", runs.len() - 3);
            }
        }
    }

    if paragraphs.len() > 5 {
        println!("\n... and {} more paragraphs", paragraphs.len() - 5);
    }

    // ===== TABLE ITERATION =====
    println!("\n\nTables:");
    println!("{}", "-".repeat(80));

    let tables = document.tables()?;
    if tables.is_empty() {
        println!("  No tables found in document");
    } else {
        for (table_idx, table) in tables.iter().enumerate() {
            let row_count = table.row_count()?;
            let col_count = table.column_count()?;

            println!(
                "\nTable #{}: {} rows × {} columns",
                table_idx + 1,
                row_count,
                col_count
            );

            let rows = table.rows()?;
            for (row_idx, row) in rows.iter().enumerate().take(3) {
                let cells = row.cells()?;
                print!("  Row #{}: ", row_idx + 1);

                for (cell_idx, cell) in cells.iter().enumerate().take(3) {
                    let cell_text = cell.text()?;
                    let display_text = cell_text.trim().replace('\n', " ");
                    let display_text = if display_text.len() > 15 {
                        format!("{}...", &display_text[..15])
                    } else {
                        display_text
                    };

                    print!("[{}] \"{}\"", cell_idx + 1, display_text);
                    if cell_idx < cells.len() - 1 && cell_idx < 2 {
                        print!(", ");
                    }
                }

                if cells.len() > 3 {
                    print!(", ... {} more cells", cells.len() - 3);
                }
                println!();
            }

            if rows.len() > 3 {
                println!("  ... and {} more rows", rows.len() - 3);
            }
        }
    }

    // ===== FULL TEXT EXTRACTION =====
    println!("\n\nFull Document Text:");
    println!("{}", "-".repeat(80));
    let text = document.text()?;
    let preview = if text.len() > 300 {
        format!("{}...\n\n(Total {} characters)", &text[..300], text.len())
    } else {
        text
    };
    println!("{}", preview);

    println!("\n{}", "=".repeat(80));
    println!("✓ Done");

    Ok(())
}
