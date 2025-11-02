//! Comprehensive demonstration of Office document CRUD operations.
//!
//! This example showcases how to create, read, and update Word documents (DOCX),
//! Excel spreadsheets (XLSX), and PowerPoint presentations (PPTX) using the Litchi library.
//!
//! Run with:
//! ```bash
//! cargo run --example office_crud_demo
//! ```

use litchi::ooxml::docx::Package as DocxPackage;
use litchi::ooxml::pptx::Package as PptxPackage;
use litchi::ooxml::xlsx::Workbook;
use litchi::sheet::WorkbookTrait;

fn main() -> Result<(), Box<dyn std::error::Error>> {
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
fn demo_docx_operations() -> Result<(), Box<dyn std::error::Error>> {
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
    pkg.properties_mut().title = Some("Product Catalog".to_string());
    pkg.properties_mut().creator = Some("Litchi Demo".to_string());
    pkg.properties_mut().description = Some("Demonstration of DOCX creation".to_string());

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
    if let Some(title) = &pkg.properties().title {
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

    pkg.properties_mut().last_modified_by = Some("Litchi Update".to_string());
    pkg.save("demo_catalog_updated.docx")?;
    println!("✓ Updated: demo_catalog_updated.docx");

    Ok(())
}

/// Demonstrate Excel spreadsheet (XLSX) operations
fn demo_xlsx_operations() -> Result<(), Box<dyn std::error::Error>> {
    println!("--- XLSX Operations ---");

    // CREATE: Create a new workbook
    println!("Creating new Excel workbook...");
    let mut wb = Workbook::create()?;

    // Add data to first worksheet
    {
        let ws = wb.worksheet_mut(0)?;
        ws.set_cell_value(1, 1, "Employee");
        ws.set_cell_value(1, 2, "Department");
        ws.set_cell_value(1, 3, "Salary");

        ws.set_cell_value(2, 1, "Alice Johnson");
        ws.set_cell_value(2, 2, "Engineering");
        ws.set_cell_value(2, 3, 85000);

        ws.set_cell_value(3, 1, "Bob Smith");
        ws.set_cell_value(3, 2, "Marketing");
        ws.set_cell_value(3, 3, 72000);

        ws.set_cell_value(4, 1, "Carol Williams");
        ws.set_cell_value(4, 2, "Sales");
        ws.set_cell_value(4, 3, 68000);

        ws.set_cell_value(5, 1, "David Brown");
        ws.set_cell_value(5, 2, "Engineering");
        ws.set_cell_value(5, 3, 92000);

        // Set freeze panes
        ws.freeze_panes(2, 1);
    }

    // Add a second worksheet for summary
    {
        let summary = wb.add_worksheet("Summary");
        summary.set_cell_value(1, 1, "Department");
        summary.set_cell_value(1, 2, "Average Salary");
        summary.set_cell_value(2, 1, "Engineering");
        summary.set_cell_formula(2, 2, "=AVERAGE(Sheet1!C2:C5)");
    }

    // Define named range
    wb.define_name("EmployeeData", "Sheet1!$A$1:$C$5");

    // Set metadata
    wb.properties_mut().title = Some("Employee Database".to_string());
    wb.properties_mut().creator = Some("Litchi Demo".to_string());

    // Save
    wb.save("demo_employees.xlsx")?;
    println!("✓ Created: demo_employees.xlsx");

    // READ: Open and read the workbook
    println!("Reading Excel workbook...");
    let wb = Workbook::open("demo_employees.xlsx")?;

    println!("  Worksheets: {}", wb.worksheet_count());
    for name in wb.worksheet_names() {
        println!("    - {}", name);
    }

    // Read data from first worksheet
    let ws = wb.worksheet_by_index(0)?;
    println!("  Sheet '{}' dimensions: {:?}", ws.name(), ws.dimensions());

    // Note: Search functionality requires the concrete Worksheet type
    // For now, we can iterate cells to search
    println!("  Worksheet loaded successfully");

    // UPDATE: Modify existing workbook
    println!("Updating Excel workbook...");
    let mut wb = Workbook::open("demo_employees.xlsx")?;

    // Add new employee
    {
        let ws = wb.worksheet_mut(0)?;
        ws.set_cell_value(6, 1, "Eve Davis");
        ws.set_cell_value(6, 2, "HR");
        ws.set_cell_value(6, 3, 65000);
    }

    wb.properties_mut().last_modified_by = Some("Litchi Update".to_string());
    wb.save("demo_employees_updated.xlsx")?;
    println!("✓ Updated: demo_employees_updated.xlsx");

    Ok(())
}

/// Demonstrate PowerPoint presentation (PPTX) operations
fn demo_pptx_operations() -> Result<(), Box<dyn std::error::Error>> {
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

    // Set metadata
    pkg.properties_mut().title = Some("Company Overview Q4 2024".to_string());
    pkg.properties_mut().creator = Some("Litchi Demo".to_string());

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

    pkg.properties_mut().last_modified_by = Some("Litchi Update".to_string());
    pkg.save("demo_presentation_updated.pptx")?;
    println!("✓ Updated: demo_presentation_updated.pptx");

    Ok(())
}
