//! Example demonstrating how to read and extract information from Office files.
//!
//! This example shows various reading operations for DOCX, XLSX, and PPTX files,
//! including text extraction, metadata reading, and content analysis.
//!
//! Run with:
//! ```bash
//! cargo run --example read_office_files
//! ```

use litchi::ooxml::api::helpers;
use litchi::ooxml::docx::Package as DocxPackage;
use litchi::ooxml::pptx::Package as PptxPackage;
use litchi::ooxml::xlsx::Workbook;
use litchi::sheet::WorkbookTrait;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== Office File Reading Demo ===\n");

    // Using the unified helper API
    demo_unified_api()?;
    println!();

    // Format-specific reading
    demo_docx_reading()?;
    println!();

    demo_xlsx_reading()?;
    println!();

    demo_pptx_reading()?;
    println!();

    Ok(())
}

/// Demonstrate unified API for reading any Office format
fn demo_unified_api() -> Result<(), Box<dyn std::error::Error>> {
    println!("--- Unified API Demo ---");

    // These helper functions work with any supported format
    let files = vec![
        "sample_document.docx",
        "sample_spreadsheet.xlsx",
        "sample_presentation.pptx",
    ];

    for file in files {
        // Note: In a real scenario, check if file exists first
        match helpers::extract_text(file) {
            Ok(text) => {
                println!("Text from {}: {} characters", file, text.len());
            },
            Err(e) => {
                println!("Could not read {}: {}", file, e);
            },
        }

        match helpers::get_properties(file) {
            Ok(props) => {
                if let Some(title) = props.title {
                    println!("  Title: {}", title);
                }
                if let Some(creator) = props.creator {
                    println!("  Creator: {}", creator);
                }
            },
            Err(e) => {
                println!("Could not read properties from {}: {}", file, e);
            },
        }
    }

    Ok(())
}

/// Demonstrate Word document reading
fn demo_docx_reading() -> Result<(), Box<dyn std::error::Error>> {
    println!("--- DOCX Reading Demo ---");

    // Check if file exists (in real usage)
    let filename = "sample_document.docx";
    println!("Analyzing: {}", filename);

    // Open document
    match DocxPackage::open(filename) {
        Ok(pkg) => {
            let doc = pkg.document()?;

            // Basic statistics
            println!("  Paragraphs: {}", doc.paragraph_count()?);
            println!("  Tables: {}", doc.table_count()?);

            // Extract all text
            let text = doc.text()?;
            println!("  Total characters: {}", text.len());
            println!("  Total words: {}", text.split_whitespace().count());

            // Analyze first few paragraphs
            let paragraphs = doc.paragraphs()?;
            println!("  First 3 paragraphs:");
            for (idx, para) in paragraphs.iter().take(3).enumerate() {
                let text = para.text()?;
                println!(
                    "    {}: {}",
                    idx + 1,
                    if text.len() > 60 {
                        format!("{}...", &text[..60])
                    } else {
                        text
                    }
                );
            }

            // Search for specific text
            let search_term = "important";
            let matches = doc.search(search_term)?;
            println!("  Found '{}' in {} paragraphs", search_term, matches.len());

            // Table analysis
            if doc.has_tables()? {
                if let Some(table) = doc.table(0)? {
                    let rows = table.rows()?;
                    println!(
                        "  First table: {}x{} (rows x cols)",
                        rows.len(),
                        rows.first()
                            .map(|r| r.cells().map(|c| c.len()).unwrap_or(0))
                            .unwrap_or(0)
                    );
                }
            }

            // Metadata
            let props = pkg.properties();
            println!("  Metadata:");
            if let Some(title) = &props.title {
                println!("    Title: {}", title);
            }
            if let Some(creator) = &props.creator {
                println!("    Creator: {}", creator);
            }
            if let Some(created) = &props.created {
                println!("    Created: {}", created);
            }
        },
        Err(e) => {
            println!("  File not found or error: {}", e);
            println!("  (This is expected if sample file doesn't exist)");
        },
    }

    Ok(())
}

/// Demonstrate Excel spreadsheet reading
fn demo_xlsx_reading() -> Result<(), Box<dyn std::error::Error>> {
    println!("--- XLSX Reading Demo ---");

    let filename = "sample_spreadsheet.xlsx";
    println!("Analyzing: {}", filename);

    match Workbook::open(filename) {
        Ok(wb) => {
            // Basic statistics
            println!("  Worksheets: {}", wb.worksheet_count());

            // List all sheets
            println!("  Sheet names:");
            for name in wb.worksheet_names() {
                println!("    - {}", name);
            }

            // Analyze first worksheet
            if let Ok(ws) = wb.worksheet_by_index(0) {
                println!("  First sheet: '{}'", ws.name());

                if let Some((min_row, min_col, max_row, max_col)) = ws.dimensions() {
                    println!(
                        "    Used range: {}x{} (rows x cols)",
                        max_row - min_row + 1,
                        max_col - min_col + 1
                    );
                }

                // Sample dimensions
                println!("    Row count: {}", ws.row_count());
                println!("    Column count: {}", ws.column_count());
            }

            // Metadata
            let props = wb.properties();
            println!("  Metadata:");
            if let Some(title) = &props.title {
                println!("    Title: {}", title);
            }
            if let Some(creator) = &props.creator {
                println!("    Creator: {}", creator);
            }
        },
        Err(e) => {
            println!("  File not found or error: {}", e);
            println!("  (This is expected if sample file doesn't exist)");
        },
    }

    Ok(())
}

/// Demonstrate PowerPoint presentation reading
fn demo_pptx_reading() -> Result<(), Box<dyn std::error::Error>> {
    println!("--- PPTX Reading Demo ---");

    let filename = "sample_presentation.pptx";
    println!("Analyzing: {}", filename);

    match PptxPackage::open(filename) {
        Ok(pkg) => {
            let pres = pkg.presentation()?;

            // Basic statistics
            println!("  Slides: {}", pres.slide_count()?);

            if let Some(width) = pres.slide_width()? {
                let inches = width as f64 / 914400.0;
                println!("  Slide width: {:.2} inches", inches);
            }

            if let Some(height) = pres.slide_height()? {
                let inches = height as f64 / 914400.0;
                println!("  Slide height: {:.2} inches", inches);
            }

            // Analyze each slide
            let slides = pres.slides()?;
            for (idx, slide) in slides.iter().enumerate() {
                println!("  Slide {}:", idx + 1);
                println!("    Shapes: {}", slide.shape_count()?);
                println!("    Has tables: {}", slide.has_tables()?);
                println!("    Has pictures: {}", slide.has_pictures()?);

                // Extract text
                let text = slide.text()?;
                let preview = if text.len() > 80 {
                    format!("{}...", &text[..80].replace('\n', " "))
                } else {
                    text.replace('\n', " ")
                };
                println!("    Text preview: {}", preview);

                // Search in this slide
                if let Ok(matches) = slide.find_text("important") {
                    if !matches.is_empty() {
                        println!("    Contains 'important' in {} shapes", matches.len());
                    }
                }
            }

            // Metadata
            let props = pkg.properties();
            println!("  Metadata:");
            if let Some(title) = &props.title {
                println!("    Title: {}", title);
            }
            if let Some(creator) = &props.creator {
                println!("    Creator: {}", creator);
            }
        },
        Err(e) => {
            println!("  File not found or error: {}", e);
            println!("  (This is expected if sample file doesn't exist)");
        },
    }

    Ok(())
}
