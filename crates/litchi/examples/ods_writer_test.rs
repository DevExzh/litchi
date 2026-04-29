//! Comprehensive ODS (OpenDocument Spreadsheet) writing example.
//!
//! This example demonstrates all writing capabilities for ODS files,
//! creating a feature-rich spreadsheet to showcase the library's capabilities.
//!
//! Run with:
//! ```bash
//! cargo run --example ods_writer_test --features odf --no-default-features
//! ```

#[cfg(feature = "odf")]
use litchi::Result;
#[cfg(feature = "odf")]
use litchi::common::Metadata;
#[cfg(feature = "odf")]
use litchi::odf::{CellValue, SpreadsheetBuilder};

#[cfg(feature = "odf")]
fn main() -> Result<()> {
    println!("=== ODS Writer Comprehensive Test ===\n");

    let output_file = "ods_writer_test_output.ods";
    println!("📝 Creating comprehensive ODS spreadsheet: {}", output_file);

    // Create a new spreadsheet using SpreadsheetBuilder
    let mut builder = SpreadsheetBuilder::new();

    // Set document metadata
    println!("✅ Setting metadata...");
    let mut metadata = Metadata::default();
    metadata.title = Some("Comprehensive ODS Writer Test Spreadsheet".to_string());
    metadata.author = Some("Litchi Library Test Suite".to_string());
    metadata.description = Some(
        "This spreadsheet demonstrates all writing capabilities of the litchi ODS writer module."
            .to_string(),
    );
    metadata.subject = Some("ODS Writer Test".to_string());
    builder.set_metadata(metadata);

    // Sheet 1: Data Types Demo
    println!("✅ Creating Sheet 1: Data Types Demo...");
    builder.add_sheet("Data Types")?;

    // Headers
    builder.add_row_with_values(&["Type", "Example", "Description"])?;

    // String values
    builder.add_row_with_cell_values(&[
        CellValue::Text("String".to_string()),
        CellValue::Text("Hello, World!".to_string()),
        CellValue::Text("Text data".to_string()),
    ])?;

    // Number values
    builder.add_row_with_cell_values(&[
        CellValue::Text("Integer".to_string()),
        CellValue::Number(42.0),
        CellValue::Text("Whole number".to_string()),
    ])?;

    builder.add_row_with_cell_values(&[
        CellValue::Text("Float".to_string()),
        CellValue::Number(3.14159),
        CellValue::Text("Decimal number".to_string()),
    ])?;

    // Boolean values
    builder.add_row_with_cell_values(&[
        CellValue::Text("Boolean (True)".to_string()),
        CellValue::Boolean(true),
        CellValue::Text("Logical true".to_string()),
    ])?;

    builder.add_row_with_cell_values(&[
        CellValue::Text("Boolean (False)".to_string()),
        CellValue::Boolean(false),
        CellValue::Text("Logical false".to_string()),
    ])?;

    // Percentage
    builder.add_row_with_cell_values(&[
        CellValue::Text("Percentage".to_string()),
        CellValue::Percentage(75.5),
        CellValue::Text("Percentage value".to_string()),
    ])?;

    // Currency
    builder.add_row_with_cell_values(&[
        CellValue::Text("Currency (USD)".to_string()),
        CellValue::Currency(1234.56, "USD".to_string()),
        CellValue::Text("US Dollars".to_string()),
    ])?;

    builder.add_row_with_cell_values(&[
        CellValue::Text("Currency (EUR)".to_string()),
        CellValue::Currency(999.99, "EUR".to_string()),
        CellValue::Text("Euros".to_string()),
    ])?;

    // Sheet 2: Sales Data with Formulas
    println!("✅ Creating Sheet 2: Sales Data with Formulas...");
    builder.add_sheet("Sales Data")?;
    builder.select_sheet(1)?;

    // Headers
    builder.add_row_with_values(&["Product", "Q1", "Q2", "Q3", "Q4", "Total"])?;

    // Product data
    builder.add_row_with_cell_values(&[
        CellValue::Text("Laptop".to_string()),
        CellValue::Number(120.0),
        CellValue::Number(135.0),
        CellValue::Number(150.0),
        CellValue::Number(145.0),
        CellValue::Empty,
    ])?;
    builder.set_cell_formula(1, 5, "of:=SUM(B2:E2)")?;

    builder.add_row_with_cell_values(&[
        CellValue::Text("Mouse".to_string()),
        CellValue::Number(450.0),
        CellValue::Number(480.0),
        CellValue::Number(520.0),
        CellValue::Number(500.0),
        CellValue::Empty,
    ])?;
    builder.set_cell_formula(2, 5, "of:=SUM(B3:E3)")?;

    builder.add_row_with_cell_values(&[
        CellValue::Text("Keyboard".to_string()),
        CellValue::Number(280.0),
        CellValue::Number(295.0),
        CellValue::Number(310.0),
        CellValue::Number(305.0),
        CellValue::Empty,
    ])?;
    builder.set_cell_formula(3, 5, "of:=SUM(B4:E4)")?;

    builder.add_row_with_cell_values(&[
        CellValue::Text("Monitor".to_string()),
        CellValue::Number(95.0),
        CellValue::Number(105.0),
        CellValue::Number(115.0),
        CellValue::Number(110.0),
        CellValue::Empty,
    ])?;
    builder.set_cell_formula(4, 5, "of:=SUM(B5:E5)")?;

    builder.add_row_with_cell_values(&[
        CellValue::Text("Headset".to_string()),
        CellValue::Number(200.0),
        CellValue::Number(220.0),
        CellValue::Number(240.0),
        CellValue::Number(230.0),
        CellValue::Empty,
    ])?;
    builder.set_cell_formula(5, 5, "of:=SUM(B6:E6)")?;

    // Totals row
    builder.add_row_with_values(&["TOTAL", "", "", "", "", ""])?;
    builder.set_cell_formula(6, 1, "of:=SUM(B2:B6)")?;
    builder.set_cell_formula(6, 2, "of:=SUM(C2:C6)")?;
    builder.set_cell_formula(6, 3, "of:=SUM(D2:D6)")?;
    builder.set_cell_formula(6, 4, "of:=SUM(E2:E6)")?;
    builder.set_cell_formula(6, 5, "of:=SUM(F2:F6)")?;

    // Sheet 3: Student Grades
    println!("✅ Creating Sheet 3: Student Grades...");
    builder.add_sheet("Student Grades")?;
    builder.select_sheet(2)?;

    // Headers
    builder.add_row_with_values(&[
        "Student Name",
        "Math",
        "Science",
        "English",
        "History",
        "Average",
    ])?;

    // Student data
    let students = vec![
        ("Alice Johnson", vec![95.0, 92.0, 88.0, 90.0]),
        ("Bob Smith", vec![78.0, 82.0, 85.0, 80.0]),
        ("Charlie Brown", vec![88.0, 90.0, 87.0, 89.0]),
        ("Diana Prince", vec![92.0, 95.0, 93.0, 94.0]),
        ("Ethan Hunt", vec![75.0, 78.0, 80.0, 77.0]),
        ("Fiona Green", vec![85.0, 88.0, 86.0, 87.0]),
    ];

    let mut row_idx = 1; // Start after header
    for (name, scores) in students {
        let mut row_values = vec![CellValue::Text(name.to_string())];
        for score in scores {
            row_values.push(CellValue::Number(score));
        }
        row_values.push(CellValue::Empty);
        builder.add_row_with_cell_values(&row_values)?;

        // Add average formula for the current row
        builder.set_cell_formula(
            row_idx,
            5,
            &format!("of:=AVERAGE(B{}:E{})", row_idx + 1, row_idx + 1),
        )?;
        row_idx += 1;
    }

    // Sheet 4: Unicode and Special Characters
    println!("✅ Creating Sheet 4: Unicode and Special Characters...");
    builder.add_sheet("Unicode Test")?;
    builder.select_sheet(3)?;

    builder.add_row_with_values(&["Language", "Greeting"])?;

    let greetings = vec![
        ("English", "Hello, World!"),
        ("Chinese", "你好，世界！"),
        ("Japanese", "こんにちは、世界！"),
        ("Korean", "안녕하세요, 세계!"),
        ("Russian", "Привет, мир!"),
        ("Arabic", "مرحبا بالعالم!"),
        ("Hebrew", "שלום, עולם!"),
        ("Greek", "Γεια σου, κόσμε!"),
        ("Emoji", "😀 🌍 🎉 🚀 ⭐"),
        ("Math", "∫ ∑ ∏ √ ∞ ≈ ≠ ≤ ≥"),
    ];

    for (lang, text) in greetings {
        builder.add_row_with_values(&[lang, text])?;
    }

    // Sheet 5: Large Data Set
    println!("✅ Creating Sheet 5: Large Data Set...");
    builder.add_sheet("Large Data")?;
    builder.select_sheet(4)?;

    builder.add_row_with_values(&["ID", "Value", "Squared", "Cubed"])?;

    for i in 1..=50 {
        builder.add_row_with_cell_values(&[
            CellValue::Number(i as f64),
            CellValue::Number(i as f64 * 10.0),
            CellValue::Empty,
            CellValue::Empty,
        ])?;

        builder.set_cell_formula(i, 2, &format!("of:=B{}*B{}", i + 1, i + 1))?;
        builder.set_cell_formula(i, 3, &format!("of:=B{}*B{}*B{}", i + 1, i + 1, i + 1))?;
    }

    // Sheet 6: Complex Layout with Mixed Data
    println!("✅ Creating Sheet 6: Complex Layout...");
    builder.add_sheet("Complex Layout")?;
    builder.select_sheet(5)?;

    // Title section
    builder.add_row_with_values(&["QUARTERLY REPORT", "", "", ""])?;
    builder.add_row_with_values(&["Year: 2024", "", "", ""])?;
    builder.add_row_with_values(&["", "", "", ""])?; // Empty row

    // Revenue section
    builder.add_row_with_values(&["REVENUE BREAKDOWN", "", "", ""])?;
    builder.add_row_with_values(&["Department", "Amount", "% of Total", ""])?;

    let departments = vec![
        ("Sales", 500000.0),
        ("Marketing", 150000.0),
        ("Engineering", 300000.0),
        ("Operations", 200000.0),
    ];

    let total: f64 = departments.iter().map(|(_, amt)| amt).sum();

    for (dept, amount) in departments {
        builder.add_row_with_cell_values(&[
            CellValue::Text(dept.to_string()),
            CellValue::Currency(amount, "USD".to_string()),
            CellValue::Percentage((amount / total) * 100.0),
            CellValue::Empty,
        ])?;
    }

    // Total row
    builder.add_row_with_cell_values(&[
        CellValue::Text("TOTAL".to_string()),
        CellValue::Currency(total, "USD".to_string()),
        CellValue::Percentage(100.0),
        CellValue::Empty,
    ])?;

    // Save the spreadsheet
    println!("💾 Saving spreadsheet to: {}", output_file);
    builder.save(output_file)?;

    println!("✅ Spreadsheet saved successfully!");
    println!("\n📊 Spreadsheet Contents:");
    println!("  - Sheets: 6");
    println!("    1. Data Types - All supported cell value types");
    println!("    2. Sales Data - Formulas (SUM)");
    println!("    3. Student Grades - Grade calculations (AVERAGE)");
    println!("    4. Unicode Test - Multilingual text");
    println!("    5. Large Data - 50 rows with formulas");
    println!("    6. Complex Layout - Report-style formatting");
    println!("  - Total Cells: ~400+");
    println!("  - Formulas: ~200+");
    println!("\n=== ODS Writer Test Complete ===");
    println!("✅ Comprehensive ODS file created successfully!");
    println!("📖 Open 'ods_writer_test_output.ods' in LibreOffice Calc to view the result.");

    Ok(())
}

#[cfg(not(feature = "odf"))]
fn main() {
    eprintln!(
        "This example requires the 'odf' feature. Try: cargo run --example ods_writer_test --features odf --no-default-features"
    );
}
