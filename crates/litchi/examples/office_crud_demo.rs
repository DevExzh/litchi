//! Comprehensive demonstration of Office document CRUD operations.
//!
//! This example showcases how to create, read, and update Word documents (DOCX),
//! Excel spreadsheets (XLSX), and PowerPoint presentations (PPTX) using the Litchi library.
//!
//! Run with:
//! ```bash
//! cargo run --example office_crud_demo
//! ```

use litchi::docx::Package as DocxPackage;
use litchi::ooxml_common::Props;
use litchi::pptx::Package as PptxPackage;
use litchi::xlsx::{Formula, Workbook};

type ExampleResult<T> = Result<T, Box<dyn std::error::Error + Send + Sync>>;

fn main() -> ExampleResult<()> {
    println!("=== Office Document CRUD Demo ===\n");

    // Demonstrate DOCX operations
    demo_docx_operations()?;
    println!();

    // Demonstrate XLSX operations
    demo_xlsx_operations()?;
    println!();

    // Demonstrate PPTX operations
    demo_pptx_operations()?;
    println!();

    println!("✓ All operations completed successfully!");

    Ok(())
}

/// Demonstrate Word document (DOCX) operations
fn demo_docx_operations() -> ExampleResult<()> {
    println!("--- DOCX Operations ---");

    // CREATE: Create a new document
    println!("Creating new Word document...");
    let mut pkg = DocxPackage::new()?;
    let doc = pkg.document_mut()?;

    // Add content
    doc.add_heading("Product Catalog", 1)?;
    doc.add_paragraph_with_text("Welcome to our comprehensive product catalog.");

    let intro_para = doc.add_paragraph();
    intro_para
        .add_run_with_text("This document contains ")
        .bold(false);
    intro_para
        .add_run_with_text("important information")
        .bold(true);
    intro_para
        .add_run_with_text(" about our products.")
        .bold(false);

    // Add a table
    doc.add_heading("Product List", 2)?;
    let table = doc.add_table(4, 3);
    table.cell(0, 0).unwrap().set_text("Product");
    table.cell(0, 1).unwrap().set_text("Price");
    table.cell(0, 2).unwrap().set_text("Stock");

    table.cell(1, 0).unwrap().set_text("Widget A");
    table.cell(1, 1).unwrap().set_text("$19.99");
    table.cell(1, 2).unwrap().set_text("150");

    table.cell(2, 0).unwrap().set_text("Widget B");
    table.cell(2, 1).unwrap().set_text("$29.99");
    table.cell(2, 2).unwrap().set_text("75");

    table.cell(3, 0).unwrap().set_text("Widget C");
    table.cell(3, 1).unwrap().set_text("$39.99");
    table.cell(3, 2).unwrap().set_text("200");

    // Set metadata
    let _ = pkg.put_props(
        Props::new()
            .title("Product Catalog")
            .creator("Litchi Demo")
            .description("Demonstration of DOCX creation"),
    );

    // Save
    pkg.save("demo_catalog.docx")?;
    println!("✓ Created: demo_catalog.docx");

    // READ: Open and read the document
    println!("Reading Word document...");
    let pkg = DocxPackage::open("demo_catalog.docx")?;
    let doc = pkg.document()?;

    println!("  Paragraphs: {}", doc.paragraph_count()?);
    println!("  Tables: {}", doc.table_count()?);

    // Search for text
    let matches = doc.search("Widget")?;
    println!("  Found 'Widget' in {} paragraphs", matches.len());

    // Access metadata
    if let Some(title) = pkg.props().and_then(|props| props.title.as_deref()) {
        println!("  Title: {}", title);
    }

    // UPDATE: Modify existing document
    println!("Updating Word document...");
    let mut pkg = DocxPackage::open("demo_catalog.docx")?;
    let doc = pkg.document_mut()?;

    doc.add_heading("Contact Information", 2)?;
    doc.add_paragraph_with_text("For more information, please contact us:");
    doc.add_paragraph_with_text("Email: sales@example.com");
    doc.add_paragraph_with_text("Phone: (555) 123-4567");

    if let Some(props) = pkg.props_mut() {
        props.last_modified_by = Some("Litchi Update".to_string());
    }
    pkg.save("demo_catalog_updated.docx")?;
    println!("✓ Updated: demo_catalog_updated.docx");

    Ok(())
}

/// Demonstrate Excel spreadsheet (XLSX) operations
fn demo_xlsx_operations() -> ExampleResult<()> {
    println!("--- XLSX Operations ---");

    // CREATE: Create a new workbook
    println!("Creating new Excel workbook...");
    let wb = Workbook::create()?;
    let mut edit = wb.edit()?;

    // XLSX authoring is an atomic semantic transaction over an immutable
    // workbook snapshot.
    {
        let mut sheet = edit
            .sheet("Sheet1")?
            .ok_or_else(|| std::io::Error::new(std::io::ErrorKind::NotFound, "Sheet1"))?;
        sheet
            .set("A1", "Employee")?
            .set("B1", "Department")?
            .set("C1", "Salary")?
            .set("A2", "Alice Johnson")?
            .set("B2", "Engineering")?
            .set("C2", 85000_i32)?
            .set("A3", "Bob Smith")?
            .set("B3", "Marketing")?
            .set("C3", 72000_i32)?
            .set("A4", "Carol Williams")?
            .set("B4", "Sales")?
            .set("C4", 68000_i32)?
            .set("A5", "David Brown")?
            .set("B5", "Engineering")?
            .set("C5", 92000_i32)?;
    }

    // Add a second worksheet for summary.
    {
        let mut summary = edit.add("Summary")?;
        summary
            .set("A1", "Department")?
            .set("B1", "Average Salary")?
            .set("A2", "Engineering")?
            .set("B2", Formula::new("AVERAGE(Sheet1!C2:C5)")?)?;
    }

    let wb = edit.commit()?.into_workbook();

    // Save
    wb.save("demo_employees.xlsx")?;
    println!("✓ Created: demo_employees.xlsx");

    // READ: Open and read the workbook
    println!("Reading Excel workbook...");
    let wb = Workbook::open("demo_employees.xlsx")?;

    println!("  Worksheets: {}", wb.len());
    for sheet in wb.sheets() {
        println!("    - {}", sheet.name());
    }

    // Read data from first worksheet
    let ws = wb
        .sheet(0)?
        .ok_or_else(|| std::io::Error::new(std::io::ErrorKind::NotFound, "first worksheet"))?;
    println!(
        "  Sheet '{}' stored extent: {:?}",
        ws.name(),
        ws.stored_extent()?
    );

    // Note: Search functionality requires the concrete Worksheet type
    // For now, we can iterate cells to search
    println!("  Worksheet loaded successfully");

    // UPDATE: Modify existing workbook
    println!("Updating Excel workbook...");
    let wb = Workbook::open("demo_employees.xlsx")?;
    let mut edit = wb.edit()?;
    let mut sheet = edit
        .sheet("Sheet1")?
        .ok_or_else(|| std::io::Error::new(std::io::ErrorKind::NotFound, "Sheet1"))?;
    sheet
        .set("A6", "Eve Davis")?
        .set("B6", "HR")?
        .set("C6", 65000_i32)?;
    let wb = edit.commit()?.into_workbook();
    wb.save("demo_employees_updated.xlsx")?;
    println!("✓ Updated: demo_employees_updated.xlsx");

    Ok(())
}

/// Demonstrate PowerPoint presentation (PPTX) operations
fn demo_pptx_operations() -> ExampleResult<()> {
    println!("--- PPTX Operations ---");

    // CREATE: Create a new presentation
    println!("Creating new PowerPoint presentation...");
    let mut pkg = PptxPackage::new()?;
    let pres = pkg.presentation_mut()?;

    // Add title slide
    let slide1 = pres.add_slide()?;
    slide1.set_title("Company Overview");
    slide1.add_text_box(
        "Q4 2024 Performance Review",
        914400,  // x: 1 inch
        2743200, // y: 3 inches
        7315200, // width: 8 inches
        914400,  // height: 1 inch
    );

    // Add agenda slide
    let slide2 = pres.add_slide()?;
    slide2.set_title("Agenda");
    slide2.add_text_box(
        "• Financial Performance\n• Product Launches\n• Team Updates\n• Future Plans",
        914400,
        2286000,
        7315200,
        2743200,
    );

    // Add content slide
    let slide3 = pres.add_slide()?;
    slide3.set_title("Financial Performance");
    slide3.add_text_box(
        "Revenue increased by 25% year-over-year",
        914400,
        2286000,
        7315200,
        914400,
    );
    slide3.add_text_box(
        "Profit margins improved to 18%",
        914400,
        3200400,
        7315200,
        914400,
    );

    // Save
    pkg.save("demo_presentation.pptx")?;
    println!("✓ Created: demo_presentation.pptx");

    // READ: Open and read the presentation
    println!("Reading PowerPoint presentation...");
    let pkg = PptxPackage::open("demo_presentation.pptx")?;
    let pres = pkg.presentation()?;

    println!("  Slides: {}", pres.slide_count()?);

    // Extract text from each slide
    for (idx, slide) in pres.slides()?.iter().enumerate() {
        let text = slide.text()?;
        if !text.is_empty() {
            println!("  Slide {}: {} shapes", idx + 1, slide.shape_count()?);
        }
    }

    // UPDATE: Modify existing presentation
    println!("Updating PowerPoint presentation...");
    let mut pkg = PptxPackage::open("demo_presentation.pptx")?;
    let pres = pkg.presentation_mut()?;

    // Add conclusion slide
    let slide4 = pres.add_slide()?;
    slide4.set_title("Conclusion");
    slide4.add_text_box(
        "Thank you for your attention!",
        914400,
        2743200,
        7315200,
        914400,
    );
    slide4.add_text_box("Questions?", 914400, 3657600, 7315200, 914400);

    pkg.save("demo_presentation_updated.pptx")?;
    println!("✓ Updated: demo_presentation_updated.pptx");

    Ok(())
}
