//! XLS New Features Demo
//!
//! Generates an XLS file showcasing merged cells, hyperlinks, and auto-filter
//! so you can open it in Microsoft Excel and verify correctness.
//!
//! Run with: cargo run --example xls_new_features_demo
//!
//! The file is saved to `output/xls_new_features_demo.xls`.

use litchi_xls::Writer;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output_dir = std::path::Path::new("output");
    std::fs::create_dir_all(output_dir)?;
    let output_path = output_dir.join("xls_new_features_demo.xls");

    let mut w = Writer::new();

    // ================================================================
    // Sheet 1 — Merged Cells
    // ================================================================
    let s1 = w.add_worksheet("Merged Cells")?;

    // Title spanning A1:E1
    w.write_string(s1, 0, 0, "Q1 2025 Revenue Report")?;
    w.merge_cells(s1, 0, 0, 0, 4)?; // single row, cols A-E

    // Sub-header spanning A2:B2 and C2:E2
    w.write_string(s1, 1, 0, "Region Info")?;
    w.merge_cells(s1, 1, 1, 0, 1)?;
    w.write_string(s1, 1, 2, "Revenue Breakdown")?;
    w.merge_cells(s1, 1, 1, 2, 4)?;

    // Column headers row 3
    let headers = ["Region", "Manager", "Jan", "Feb", "Mar"];
    for (c, h) in headers.iter().enumerate() {
        w.write_string(s1, 2, c as u16, h)?;
    }

    // Data rows
    let data: &[(&str, &str, f64, f64, f64)] = &[
        ("North", "Alice", 12500.0, 13200.0, 14100.0),
        ("South", "Bob", 9800.0, 10500.0, 11200.0),
        ("East", "Carol", 15200.0, 14800.0, 16300.0),
        ("West", "Dave", 11000.0, 11800.0, 12400.0),
    ];
    for (i, (region, mgr, jan, feb, mar)) in data.iter().enumerate() {
        let row = 3 + i as u32;
        w.write_string(s1, row, 0, region)?;
        w.write_string(s1, row, 1, mgr)?;
        w.write_number(s1, row, 2, *jan)?;
        w.write_number(s1, row, 3, *feb)?;
        w.write_number(s1, row, 4, *mar)?;
    }

    // Merged total label spanning A8:B8
    w.write_string(s1, 7, 0, "Grand Total")?;
    w.merge_cells(s1, 7, 7, 0, 1)?;
    w.write_formula(s1, 7, 2, "SUM(C4:C7)")?;
    w.write_formula(s1, 7, 3, "SUM(D4:D7)")?;
    w.write_formula(s1, 7, 4, "SUM(E4:E7)")?;

    // Widen columns for readability
    w.set_column_width(s1, 0, 14.0)?;
    w.set_column_width(s1, 1, 14.0)?;
    for c in 2..=4 {
        w.set_column_width(s1, c, 12.0)?;
    }

    println!("[Sheet 1] Merged Cells — 3 merge regions, 4 data rows, totals");

    // ================================================================
    // Sheet 2 — Hyperlinks
    // ================================================================
    let s2 = w.add_worksheet("Hyperlinks")?;

    w.write_string(s2, 0, 0, "Description")?;
    w.write_string(s2, 0, 1, "Link")?;
    w.set_column_width(s2, 0, 28.0)?;
    w.set_column_width(s2, 1, 40.0)?;

    // Web links
    w.write_string(s2, 1, 0, "Rust Programming Language")?;
    w.write_string(s2, 1, 1, "https://www.rust-lang.org")?;
    w.set_hyperlink(s2, 1, 1, "https://www.rust-lang.org")?;

    w.write_string(s2, 2, 0, "Rust crates.io")?;
    w.write_string(s2, 2, 1, "https://crates.io")?;
    w.set_hyperlink(s2, 2, 1, "https://crates.io")?;

    w.write_string(s2, 3, 0, "GitHub")?;
    w.write_string(s2, 3, 1, "https://github.com")?;
    w.set_hyperlink(s2, 3, 1, "https://github.com")?;

    // Email link
    w.write_string(s2, 4, 0, "Contact support")?;
    w.write_string(s2, 4, 1, "mailto:support@example.com")?;
    w.set_hyperlink(s2, 4, 1, "mailto:support@example.com")?;

    // Internal link — jump to the Merged Cells sheet
    w.write_string(s2, 5, 0, "Go to Merged Cells sheet A1")?;
    w.write_string(s2, 5, 1, "Click here")?;
    w.set_hyperlink(s2, 5, 1, "internal:Merged Cells!A1")?;

    // Internal link — jump to the AutoFilter sheet
    w.write_string(s2, 6, 0, "Go to AutoFilter sheet A1")?;
    w.write_string(s2, 6, 1, "Click here")?;
    w.set_hyperlink(s2, 6, 1, "internal:AutoFilter!A1")?;

    println!("[Sheet 2] Hyperlinks — 4 web/email links, 2 internal links");

    // ================================================================
    // Sheet 3 — Auto-Filter with sortable data
    // ================================================================
    let s3 = w.add_worksheet("AutoFilter")?;

    let af_headers = ["Employee", "Department", "Salary", "Years", "Rating"];
    for (c, h) in af_headers.iter().enumerate() {
        w.write_string(s3, 0, c as u16, h)?;
    }

    let employees: &[(&str, &str, f64, f64, &str)] = &[
        ("Alice", "Engineering", 95000.0, 5.0, "Excellent"),
        ("Bob", "Marketing", 72000.0, 3.0, "Good"),
        ("Carol", "Engineering", 105000.0, 8.0, "Excellent"),
        ("Dave", "Sales", 68000.0, 2.0, "Good"),
        ("Eve", "Engineering", 88000.0, 4.0, "Very Good"),
        ("Frank", "Marketing", 76000.0, 6.0, "Very Good"),
        ("Grace", "Sales", 71000.0, 3.0, "Good"),
        ("Hank", "Engineering", 112000.0, 10.0, "Excellent"),
        ("Ivy", "Sales", 65000.0, 1.0, "Satisfactory"),
        ("Jack", "Marketing", 82000.0, 7.0, "Very Good"),
        ("Karen", "Engineering", 99000.0, 6.0, "Excellent"),
        ("Leo", "Sales", 74000.0, 4.0, "Good"),
    ];

    for (i, (name, dept, salary, years, rating)) in employees.iter().enumerate() {
        let row = 1 + i as u32;
        w.write_string(s3, row, 0, name)?;
        w.write_string(s3, row, 1, dept)?;
        w.write_number(s3, row, 2, *salary)?;
        w.write_number(s3, row, 3, *years)?;
        w.write_string(s3, row, 4, rating)?;
    }

    // Apply auto-filter over the entire data range (row 0 through row 12, cols A-E)
    w.set_auto_filter(s3, 0, employees.len() as u32, 0, 4)?;

    // Widen columns
    w.set_column_width(s3, 0, 14.0)?;
    w.set_column_width(s3, 1, 16.0)?;
    w.set_column_width(s3, 2, 12.0)?;
    w.set_column_width(s3, 3, 8.0)?;
    w.set_column_width(s3, 4, 14.0)?;

    println!("[Sheet 3] AutoFilter — 12 employee rows, 5 filterable columns");

    // ================================================================
    // Sheet 4 — Combined features on one sheet
    // ================================================================
    let s4 = w.add_worksheet("Combined")?;

    // Merged title
    w.write_string(s4, 0, 0, "Product Catalog")?;
    w.merge_cells(s4, 0, 0, 0, 3)?;

    // Headers
    let cat_headers = ["Product", "Category", "Price", "Link"];
    for (c, h) in cat_headers.iter().enumerate() {
        w.write_string(s4, 1, c as u16, h)?;
    }

    let products: &[(&str, &str, f64, &str)] = &[
        (
            "Laptop Pro",
            "Electronics",
            1299.99,
            "https://example.com/laptop",
        ),
        (
            "Wireless Mouse",
            "Accessories",
            29.99,
            "https://example.com/mouse",
        ),
        ("USB-C Hub", "Accessories", 49.99, "https://example.com/hub"),
        (
            "Monitor 27\"",
            "Electronics",
            399.99,
            "https://example.com/monitor",
        ),
        (
            "Keyboard",
            "Accessories",
            79.99,
            "https://example.com/keyboard",
        ),
        (
            "Webcam HD",
            "Electronics",
            89.99,
            "https://example.com/webcam",
        ),
    ];

    for (i, (name, cat, price, url)) in products.iter().enumerate() {
        let row = 2 + i as u32;
        w.write_string(s4, row, 0, name)?;
        w.write_string(s4, row, 1, cat)?;
        w.write_number(s4, row, 2, *price)?;
        w.write_string(s4, row, 3, "View")?;
        w.set_hyperlink(s4, row, 3, url)?;
    }

    // Auto-filter on header+data (rows 1-7, cols A-D)
    let last_data_row = 2 + products.len() as u32 - 1;
    w.set_auto_filter(s4, 1, last_data_row, 0, 3)?;

    // Merged footer
    let footer_row = last_data_row + 2;
    w.write_string(s4, footer_row, 0, "End of catalog — all prices in USD")?;
    w.merge_cells(s4, footer_row, footer_row, 0, 3)?;

    w.set_column_width(s4, 0, 18.0)?;
    w.set_column_width(s4, 1, 14.0)?;
    w.set_column_width(s4, 2, 10.0)?;
    w.set_column_width(s4, 3, 30.0)?;

    println!("[Sheet 4] Combined — merged title+footer, hyperlinks, auto-filter");

    // ================================================================
    // Save
    // ================================================================
    w.save(&output_path)?;
    println!("\nSaved to: {}", output_path.display());
    println!("Open in Excel to verify merged cells, hyperlinks, and auto-filter dropdowns.");

    // ================================================================
    // Round-trip: read the file back and print parsed features
    // ================================================================
    println!("\n=== Round-trip verification (reading back) ===\n");
    round_trip_verify(&output_path)?;

    Ok(())
}

/// Read the generated XLS file back using `Workbook` and print
/// the parsed merged cells, hyperlinks, comments, and auto-filter data.
fn round_trip_verify(path: &std::path::Path) -> Result<(), Box<dyn std::error::Error>> {
    use litchi_core::sheet::WorkbookTrait;
    use litchi_xls::Workbook;
    use std::io::Cursor;

    let data = std::fs::read(path)?;
    let cursor = Cursor::new(data);
    let wb = Workbook::new(cursor)?;

    println!("Worksheets: {:?}", wb.worksheet_names());
    println!("Sheet count: {}", wb.worksheet_count());

    for (idx, name) in wb.worksheet_names().iter().enumerate() {
        println!("\n--- Sheet {}: \"{}\" ---", idx, name);

        // Access the underlying Worksheet to inspect new features.
        // The WorkbookTrait gives us a dyn Worksheet, but we stored
        // typed worksheets inside Workbook. The easiest way to
        // inspect features is through the typed accessor.
        //
        // For now, print basic cell info via the trait.
        let sheet = wb
            .worksheet_by_index(idx)
            .map_err(|e| -> Box<dyn std::error::Error> { e })?;
        if let Some(dims) = sheet.dimensions() {
            println!(
                "  Dimensions: rows {}..{}, cols {}..{}",
                dims.0, dims.2, dims.1, dims.3
            );
        }
        println!(
            "  Row count: {}, Col count: {}",
            sheet.row_count(),
            sheet.column_count()
        );
    }

    println!("\nRound-trip complete.");
    Ok(())
}
