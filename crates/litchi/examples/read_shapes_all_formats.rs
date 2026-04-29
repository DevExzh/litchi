//! Read and extract shapes from DOC, PPT, and XLS files
//!
//! This example demonstrates how to read Escher shapes from all three
//! legacy Office formats using the shared Escher module.
//!
//! Run with: cargo run --example read_shapes_all_formats <file.doc|file.ppt|file.xls>

use litchi::ole::doc::Package as DocPackage;
use litchi::ole::ppt::Package as PptPackage;
use litchi::ole::xls::XlsWorkbook;
use litchi::sheet::WorkbookTrait;
use std::env;
use std::error::Error;
use std::fs::File;
use std::path::Path;

fn main() -> Result<(), Box<dyn Error>> {
    let args: Vec<String> = env::args().collect();
    if args.len() != 2 {
        eprintln!("Usage: {} <file.doc|file.ppt|file.xls>", args[0]);
        eprintln!();
        eprintln!("Example:");
        eprintln!("  {} document.doc", args[0]);
        eprintln!("  {} presentation.ppt", args[0]);
        eprintln!("  {} workbook.xls", args[0]);
        std::process::exit(1);
    }

    let file_path = &args[1];
    let path = Path::new(file_path);

    if !path.exists() {
        eprintln!("Error: File not found: {}", file_path);
        std::process::exit(1);
    }

    // Detect file type by extension
    match path.extension().and_then(|s| s.to_str()) {
        Some("doc") => read_doc_shapes(file_path)?,
        Some("ppt") => read_ppt_shapes(file_path)?,
        Some("xls") => read_xls_shapes(file_path)?,
        _ => {
            eprintln!("Error: Unsupported file type. Expected .doc, .ppt, or .xls");
            std::process::exit(1);
        },
    }

    Ok(())
}

fn read_doc_shapes(file_path: &str) -> Result<(), Box<dyn Error>> {
    println!("Reading DOC file: {}", file_path);
    println!("================");
    println!();

    let mut pkg = DocPackage::open(file_path)?;
    let ole = pkg.ole_file();

    // Use the DOC shapes module
    use litchi::ole::doc::shapes;

    let shape_count = shapes::count_shapes(ole)?;
    println!("Total shapes: {}", shape_count);

    if shape_count > 0 {
        let extracted_shapes = shapes::extract_shapes(ole)?;
        println!();
        for (i, shape) in extracted_shapes.iter().enumerate() {
            println!("Shape #{}: {:?}", i + 1, shape.shape_type);
            println!("  ID: {}", shape.shape_id);
            if let Some(ref text) = shape.text {
                println!("  Text: {}", text);
            }
            if shape.is_group {
                println!("  Group with {} children", shape.children.len());
            }
        }

        // Extract text from shapes
        println!();
        let shape_text = shapes::extract_shape_text(ole)?;
        if !shape_text.is_empty() {
            println!("All shape text:\n{}", shape_text);
        }
    } else {
        println!("No shapes found in document.");
    }

    Ok(())
}

fn read_ppt_shapes(file_path: &str) -> Result<(), Box<dyn Error>> {
    println!("Reading PPT file: {}", file_path);
    println!("================");
    println!();

    let mut pkg = PptPackage::open(file_path)?;
    let ppt = pkg.presentation()?;

    println!("Total slides: {}", ppt.slide_count());
    println!();

    for (slide_idx, slide) in ppt.slides()?.into_iter().enumerate() {
        println!("\nSlide {}:", slide_idx + 1);

        let shapes = slide.shapes()?;
        println!("  Shapes: {}", shapes.len());

        for (shape_idx, shape) in shapes.iter().enumerate() {
            println!("    Shape #{}: {:?}", shape_idx + 1, shape.shape_type());
            if let Ok(text) = shape.text()
                && !text.is_empty()
            {
                println!("      Text: {}", text.lines().next().unwrap_or(&text));
            }
        }
        println!();
    }

    Ok(())
}

fn read_xls_shapes(file_path: &str) -> Result<(), Box<dyn Error>> {
    println!("Reading XLS file: {}", file_path);
    println!("================");
    println!();

    let file = File::open(file_path)?;
    let workbook = XlsWorkbook::new(file)?;

    println!("Worksheets: {}", workbook.worksheet_names().len());
    println!();

    // Note: XLS shape extraction requires access to the raw workbook stream
    // which is not directly exposed by XlsWorkbook API
    println!("XLS shape extraction requires raw workbook data access.");
    println!("Use litchi::ole::xls::shapes::extract_shapes_from_workbook() with raw data.");
    println!();
    println!("Example:");
    println!("  let ole_file = OleFile::open(file)?;");
    println!("  let workbook_data = ole_file.open_stream(&[\"Workbook\"])?;");
    println!("  let shapes = extract_shapes_from_workbook(&workbook_data)?;");

    Ok(())
}
