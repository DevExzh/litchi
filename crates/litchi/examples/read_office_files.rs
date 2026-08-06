//! Example demonstrating how to read and extract information from Office files.
//!
//! This example shows the standalone DOCX, XLSX, and PPTX package readers.
//!
//! Run with:
//! ```bash
//! cargo run --example read_office_files
//! ```

use litchi::pptx::shape::Shape;
use litchi::xlsx::Package as XlsxPackage;
use litchi::{docx::Package as DocxPackage, pptx::Package as PptxPackage};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== Office File Reading Demo ===\n");

    // Read each supported format through its standalone package facade.
    demo_unified_api()?;
    println!();

    demo_docx_reading()?;
    println!();

    demo_xlsx_reading()?;
    println!();

    demo_pptx_reading()?;
    println!();

    Ok(())
}

/// Demonstrate format dispatch without the removed convenience helper API.
fn demo_unified_api() -> Result<(), Box<dyn std::error::Error>> {
    println!("--- Unified API Demo ---");

    let files = [
        "sample_document.docx",
        "sample_spreadsheet.xlsx",
        "sample_presentation.pptx",
    ];

    for file in files {
        match std::path::Path::new(file)
            .extension()
            .and_then(|extension| extension.to_str())
        {
            Some("docx") => match DocxPackage::open(file) {
                Ok(package) => {
                    let document = package.document()?;
                    let text = document.text()?;
                    println!("Text from {file}: {} characters", text.len());
                    print_props("  ", package.props());
                },
                Err(error) => println!("Could not read {file}: {error}"),
            },
            Some("xlsx") => match XlsxPackage::open(file) {
                Ok(package) => {
                    let workbook = package.workbook()?;
                    let names: Vec<_> = workbook
                        .sheets()
                        .map(|sheet| sheet.name().to_owned())
                        .collect();
                    println!("Worksheets from {file}: {}", names.join(", "));
                },
                Err(error) => println!("Could not read {file}: {error}"),
            },
            Some("pptx") => match PptxPackage::open(file) {
                Ok(package) => {
                    let presentation = package.presentation()?;
                    let text = presentation.text()?;
                    println!("Text from {file}: {} characters", text.len());
                    print_pptx_props("  ", &package)?;
                },
                Err(error) => println!("Could not read {file}: {error}"),
            },
            _ => println!("Unsupported Office file: {file}"),
        }
    }

    Ok(())
}

fn print_props(prefix: &str, props: Option<&litchi::ooxml_common::Props>) {
    if let Some(props) = props {
        if let Some(title) = &props.title {
            println!("{prefix}Title: {title}");
        }
        if let Some(creator) = &props.creator {
            println!("{prefix}Creator: {creator}");
        }
    } else {
        println!("{prefix}No core properties");
    }
}

fn print_pptx_props(prefix: &str, package: &PptxPackage) -> Result<(), Box<dyn std::error::Error>> {
    let props = litchi::ooxml_common::properties::read(package.opc()?)
        .map_err(|error| format!("could not read PPTX core properties: {error}"))?;
    print_props(prefix, props.as_ref());
    Ok(())
}

/// Demonstrate WordprocessingML reading and core-property access.
fn demo_docx_reading() -> Result<(), Box<dyn std::error::Error>> {
    println!("--- DOCX Reading Demo ---");

    let filename = "sample_document.docx";
    println!("Analyzing: {filename}");

    match DocxPackage::open(filename) {
        Ok(package) => {
            let document = package.document()?;

            println!("  Paragraphs: {}", document.paragraph_count()?);
            println!("  Tables: {}", document.table_count()?);

            let text = document.text()?;
            println!("  Total characters: {}", text.len());
            println!("  Total words: {}", text.split_whitespace().count());

            println!("  First 3 paragraphs:");
            for (index, paragraph) in document.paragraphs()?.iter().take(3).enumerate() {
                let text = paragraph.text()?;
                println!(
                    "    {}: {}",
                    index + 1,
                    text.get(..60)
                        .map_or_else(|| format!("{text}..."), |preview| preview.to_owned())
                );
            }

            let matches = document.search("important")?;
            println!("  Found 'important' in {} paragraphs", matches.len());

            if document.has_tables()?
                && let Some(table) = document.table(0)?
            {
                let rows = table.rows()?;
                println!(
                    "  First table: {}x{} (rows x cols)",
                    rows.len(),
                    rows.first()
                        .map(|row| row.cells().map(|cells| cells.len()).unwrap_or(0))
                        .unwrap_or(0)
                );
            }

            println!("  Metadata:");
            print_props("    ", package.props());
        },
        Err(error) => {
            println!("  File not found or error: {error}");
            println!("  (This is expected if the sample file does not exist)");
        },
    }

    Ok(())
}

/// Demonstrate SpreadsheetML package, workbook, and worksheet access.
fn demo_xlsx_reading() -> Result<(), Box<dyn std::error::Error>> {
    println!("--- XLSX Reading Demo ---");

    let filename = "sample_spreadsheet.xlsx";
    println!("Analyzing: {filename}");

    match XlsxPackage::open(filename) {
        Ok(package) => {
            let workbook = package.workbook()?;
            println!("  Worksheets: {}", workbook.len());

            println!("  Sheet names:");
            for sheet in workbook.sheets() {
                println!("    - {}", sheet.name());
            }

            if let Some(sheet) = workbook.sheet(0usize)? {
                println!("  First sheet: '{}'", sheet.name());
                println!("    Stored rows: {}", sheet.rows()?.len());

                // The worksheet reader is sparse. Scanning the checked Excel
                // grid counts physical cells without materializing empty ones.
                let stored_cells = sheet.cells("A1:XFD1048576")?.count();
                println!("    Stored cells: {stored_cells}");

                if let Some((address, cell)) = sheet.cells("A1:XFD1048576")?.next() {
                    println!("    First cell: {:?} = {:?}", address, cell);
                }
            }
        },
        Err(error) => {
            println!("  File not found or error: {error}");
            println!("  (This is expected if the sample file does not exist)");
        },
    }

    Ok(())
}

/// Demonstrate PresentationML slide, shape, text, and metadata access.
fn demo_pptx_reading() -> Result<(), Box<dyn std::error::Error>> {
    println!("--- PPTX Reading Demo ---");

    let filename = "sample_presentation.pptx";
    println!("Analyzing: {filename}");

    match PptxPackage::open(filename) {
        Ok(package) => {
            let presentation = package.presentation()?;
            let (width, height) = presentation.slide_size()?;
            println!("  Slides: {}", presentation.slide_count()?);
            println!(
                "  Slide size: {:.2} x {:.2} inches",
                emu_to_inches(width),
                emu_to_inches(height)
            );

            for (index, slide) in presentation.slides()?.into_iter().enumerate() {
                let scene = slide.shapes()?;
                let table_count = scene
                    .iter()
                    .filter(|shape| matches!(shape, Shape::Table(_)))
                    .count();
                let picture_count = scene
                    .iter()
                    .filter(|shape| matches!(shape, Shape::Picture(_)))
                    .count();
                let text = slide.text()?;

                println!("  Slide {}:", index + 1);
                println!("    Shapes: {}", scene.len());
                println!("    Tables: {table_count}");
                println!("    Pictures: {picture_count}");
                println!("    Text preview: {}", preview(&text, 80));

                if text.contains("important") {
                    println!("    Contains 'important'");
                }
            }

            println!("  Metadata:");
            print_pptx_props("    ", &package)?;
        },
        Err(error) => {
            println!("  File not found or error: {error}");
            println!("  (This is expected if the sample file does not exist)");
        },
    }

    Ok(())
}

fn emu_to_inches(value: i64) -> f64 {
    value as f64 / 914_400.0
}

fn preview(text: &str, limit: usize) -> String {
    let flattened = text.replace('\n', " ");
    if flattened.len() > limit {
        format!("{}...", &flattened[..limit])
    } else {
        flattened
    }
}
