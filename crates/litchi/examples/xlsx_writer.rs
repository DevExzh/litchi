/// Comprehensive example demonstrating XLSX writing features.
///
/// This example shows how to:
/// - Create a new XLSX workbook
/// - Add worksheets with data and formulas
/// - Apply cell formatting (fonts, fills, borders)
/// - Set column widths and row heights
/// - Merge cells
/// - Add data validation
/// - Configure freeze panes
/// - Save the workbook
///
/// Usage: cargo run --example xlsx_writer -- <output-file.xlsx>
use litchi::ooxml::xlsx::styles::alignment::{Horizontal, Vertical};
use litchi::ooxml::xlsx::{
    Alignment, CellFill, CellFillPatternType, CellFont, CellFormat, Workbook,
};
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    // Get output filename from command line arguments
    let args: Vec<String> = env::args().collect();
    let output_file = if args.len() >= 2 {
        &args[1]
    } else {
        "output.xlsx"
    };

    println!("📊 Creating comprehensive XLSX workbook...");
    println!("{}", "=".repeat(80));

    // Create a new workbook
    let mut workbook = Workbook::create()?;

    // ============================================================================
    // Sheet 1: Sales Data with Formatting
    // ============================================================================
    println!("\n📋 Creating 'Sales Data' sheet...");
    let sheet1 = workbook.add_worksheet("Sales Data");

    // Header row
    sheet1.set_cell_value(0, 0, "Product");
    sheet1.set_cell_value(0, 1, "Quantity");
    sheet1.set_cell_value(0, 2, "Price");
    sheet1.set_cell_value(0, 3, "Total");

    // Data rows
    sheet1.set_cell_value(1, 0, "Laptops");
    sheet1.set_cell_value(1, 1, 150.0);
    sheet1.set_cell_value(1, 2, 999.99);
    sheet1.set_cell_formula(1, 3, "B2*C2");

    sheet1.set_cell_value(2, 0, "Mice");
    sheet1.set_cell_value(2, 1, 500.0);
    sheet1.set_cell_value(2, 2, 25.50);
    sheet1.set_cell_formula(2, 3, "B3*C3");

    sheet1.set_cell_value(3, 0, "Keyboards");
    sheet1.set_cell_value(3, 1, 300.0);
    sheet1.set_cell_value(3, 2, 79.99);
    sheet1.set_cell_formula(3, 3, "B4*C4");

    // Total row
    sheet1.set_cell_value(4, 0, "TOTAL");
    sheet1.set_cell_formula(4, 3, "SUM(D2:D4)");

    // Format header row (bold, gray background)
    println!("   Applying header formatting...");
    let header_format = CellFormat {
        font: Some(CellFont {
            name: Some("Calibri".to_string()),
            size: Some(12.0),
            bold: true,
            italic: false,
            underline: false,
            color: None,
        }),
        fill: Some(CellFill {
            pattern_type: CellFillPatternType::Solid,
            fg_color: Some("FFD3D3D3".to_string()), // Light gray
            bg_color: None,
        }),
        ..CellFormat::default()
    };

    for col in 0..4 {
        sheet1.set_cell_format(0, col, header_format.clone());
    }

    // Format total row (bold)
    println!("   Applying total row formatting...");
    let total_format = CellFormat {
        font: Some(CellFont {
            name: Some("Calibri".to_string()),
            size: Some(11.0),
            bold: true,
            italic: false,
            underline: false,
            color: None,
        }),
        ..CellFormat::default()
    };

    sheet1.set_cell_format(4, 0, total_format.clone());
    sheet1.set_cell_format(4, 3, total_format);

    // Set column widths
    println!("   Setting column widths...");
    sheet1.set_column_width(0, 15.0); // Product column
    sheet1.set_column_width(1, 10.0); // Quantity column
    sheet1.set_column_width(2, 10.0); // Price column
    sheet1.set_column_width(3, 12.0); // Total column

    // Set header row height
    sheet1.set_row_height(0, 20.0);

    // Freeze the header row
    println!("   Freezing panes...");
    sheet1.freeze_panes(1, 0);

    // ============================================================================
    // Sheet 2: Report with Merged Cells
    // ============================================================================
    println!("\n📑 Creating 'Quarterly Report' sheet...");
    let sheet2 = workbook.add_worksheet("Quarterly Report");

    // Title (merged cell)
    sheet2.set_cell_value(0, 0, "Q4 2024 Sales Report");
    sheet2.merge_cells(0, 0, 0, 3); // Merge A1:D1

    let title_format = CellFormat {
        font: Some(CellFont {
            name: Some("Arial".to_string()),
            size: Some(16.0),
            bold: true,
            italic: false,
            underline: false,
            color: Some("FF0000FF".to_string()), // Blue
        }),
        alignment: Some(Alignment::both(Horizontal::Center, Vertical::Center)),
        ..CellFormat::default()
    };
    sheet2.set_cell_format(0, 0, title_format);

    // TODO: Add hyperlink support
    // sheet2.set_cell_value(2, 0, "Visit our website");
    // sheet2.set_hyperlink(2, 0, "https://www.example.com");

    // TODO: Add cell comment support
    // sheet2.set_cell_comment(2, 0, "Click to visit our website");

    // ============================================================================
    // Sheet 3: Data Validation
    // ============================================================================
    println!("\n✅ Creating 'Input Form' sheet with validation...");
    let sheet3 = workbook.add_worksheet("Input Form");

    sheet3.set_cell_value(0, 0, "Employee Name:");
    sheet3.set_cell_value(1, 0, "Department:");
    sheet3.set_cell_value(2, 0, "Status:");
    sheet3.set_cell_value(3, 0, "Rating (1-5):");

    // Add data validation for department (list)
    use litchi::ooxml::xlsx::{DataValidationOperator, DataValidationType};

    println!("   Adding data validation...");
    sheet3.add_data_validation(
        "B2",
        DataValidationType::List {
            values: vec![
                "Engineering".to_string(),
                "Sales".to_string(),
                "Marketing".to_string(),
                "HR".to_string(),
            ],
        },
        false,
        None,
        None,
        false,
        None,
        None,
    );

    // Add data validation for status (list)
    sheet3.add_data_validation(
        "B3",
        DataValidationType::List {
            values: vec![
                "Active".to_string(),
                "Inactive".to_string(),
                "On Leave".to_string(),
            ],
        },
        true,
        Some("Select Status"),
        Some("Please select a valid status"),
        true,
        Some("Invalid Input"),
        Some("Please select from the list"),
    );

    // Add data validation for rating (whole number 1-5)
    sheet3.add_data_validation(
        "B4",
        DataValidationType::Whole {
            operator: DataValidationOperator::Between,
            value1: 1,
            value2: Some(5),
        },
        true,
        Some("Enter Rating"),
        Some("Enter a rating from 1 to 5"),
        true,
        Some("Invalid Rating"),
        Some("Rating must be between 1 and 5"),
    );

    // ============================================================================
    // Sheet 4: Chart Data (Charts are partially supported)
    // ============================================================================
    println!("\n📈 Creating 'Chart Data' sheet...");
    let sheet4 = workbook.add_worksheet("Chart Data");

    sheet4.set_cell_value(0, 0, "Month");
    sheet4.set_cell_value(0, 1, "Sales");

    let months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
    let sales = [12000.0, 15000.0, 18000.0, 16000.0, 21000.0, 24000.0];

    for (i, (month, sale)) in months.iter().zip(sales.iter()).enumerate() {
        let row = (i + 1) as u32;
        sheet4.set_cell_value(row, 0, *month);
        sheet4.set_cell_value(row, 1, *sale);
    }

    // TODO: Add chart support
    // sheet4.add_chart(
    //     ChartType::Column,
    //     "Monthly Sales",
    //     "A1:B7",
    //     (8, 0, 20, 6),
    //     true,
    // );

    // ============================================================================
    // Named Ranges
    // ============================================================================
    println!("\n🏷️  Defining named ranges...");
    workbook.define_name("SalesData", "Sheet1!$A$2:$D$4");
    workbook.define_name("TotalSales", "Sheet1!$D$5");

    // ============================================================================
    // Save the workbook
    // ============================================================================
    println!("\n💾 Saving workbook to: {}", output_file);
    workbook.save(output_file)?;

    println!("\n{}", "=".repeat(80));
    println!("✅ Workbook created successfully!");
    println!("{}", "=".repeat(80));
    println!("\n📝 The workbook contains:");
    println!("   • 'Sales Data' - Formatted data with formulas and frozen header");
    println!("   • 'Quarterly Report' - Merged cells with title formatting");
    println!("   • 'Input Form' - Data validation rules for controlled input");
    println!("   • 'Chart Data' - Sample data for charting");
    println!("\n💡 Features demonstrated:");
    println!("   ✓ Cell values (text, numbers, formulas)");
    println!("   ✓ Cell formatting (fonts, fills, borders)");
    println!("   ✓ Column widths and row heights");
    println!("   ✓ Merged cells");
    println!("   ✓ Data validation");
    println!("   ✓ Freeze panes");
    println!("   ✓ Named ranges");
    println!("\n🚧 Features not yet implemented (see TODOs in code):");
    println!("   • Hyperlinks");
    println!("   • Cell comments");
    println!("   • Charts");
    println!("   • Conditional formatting");
    println!("   • Page setup (headers, footers, print area)");
    println!(
        "\n📖 Open '{}' with Excel or LibreOffice to view the results.",
        output_file
    );

    Ok(())
}
