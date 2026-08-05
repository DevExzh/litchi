//! AutoFilter Conditions & Sort — XLS Writer Example
//!
//! Demonstrates the new `add_filter_condition` and `set_sort` APIs on `Writer`.
//! Generates an XLS file with:
//!
//! - A data table of products with prices and stock levels.
//! - An AutoFilter range with DOPER conditions on the "Price" column (> 50).
//! - A sort configuration (primary key: Price descending).
//!
//! Run with: `cargo run --example xls_autofilter_conditions`
//!
//! The file is saved to `output/xls_autofilter_conditions.xls`.

use litchi_xls::Writer;
use litchi_xls::writer::AutoFilterConditionWrite;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output_dir = std::path::Path::new("output");
    std::fs::create_dir_all(output_dir)?;
    let output_path = output_dir.join("xls_autofilter_conditions.xls");

    let mut w = Writer::new();

    // ================================================================
    // Sheet 1 — AutoFilter with numeric conditions
    // ================================================================
    let s = w.add_worksheet("Filtered Products")?;

    // Headers (row 0)
    let headers = ["Product", "Category", "Price", "Stock", "Rating"];
    for (c, h) in headers.iter().enumerate() {
        w.write_string(s, 0, c as u16, h)?;
    }

    // Data rows
    let products: &[(&str, &str, f64, f64, &str)] = &[
        ("Widget A", "Electronics", 149.99, 42.0, "Excellent"),
        ("Widget B", "Clothing", 29.99, 120.0, "Good"),
        ("Widget C", "Electronics", 79.50, 85.0, "Very Good"),
        ("Widget D", "Home", 12.00, 300.0, "Satisfactory"),
        ("Widget E", "Electronics", 249.00, 18.0, "Excellent"),
        ("Widget F", "Clothing", 45.00, 200.0, "Good"),
        ("Widget G", "Home", 89.99, 55.0, "Very Good"),
        ("Widget H", "Electronics", 199.00, 30.0, "Excellent"),
        ("Widget I", "Clothing", 55.00, 95.0, "Good"),
        ("Widget J", "Home", 34.99, 150.0, "Satisfactory"),
    ];

    for (i, (name, cat, price, stock, rating)) in products.iter().enumerate() {
        let row = 1 + i as u32;
        w.write_string(s, row, 0, name)?;
        w.write_string(s, row, 1, cat)?;
        w.write_number(s, row, 2, *price)?;
        w.write_number(s, row, 3, *stock)?;
        w.write_string(s, row, 4, rating)?;
    }

    let last_row = products.len() as u32;

    // Apply AutoFilter over the entire data range.
    w.set_auto_filter(s, 0, last_row, 0, 4)?;

    // Add a filter condition on column 2 (Price): keep rows where Price > 50.
    // Operator 0x06 = "greater than" in BIFF8 DOPER encoding.
    w.add_filter_condition(
        s,
        2, // column index within filter range (0-based)
        false,
        AutoFilterConditionWrite::Number {
            operator: 0x06, // >
            value: 50.0,
        },
        AutoFilterConditionWrite::None,
    )?;

    // Set a sort configuration: sort by Price (column 2) descending.
    w.set_sort(s, false, false, &[(2, true)])?;

    // Column widths for readability.
    w.set_column_width(s, 0, 14.0)?;
    w.set_column_width(s, 1, 14.0)?;
    w.set_column_width(s, 2, 10.0)?;
    w.set_column_width(s, 3, 10.0)?;
    w.set_column_width(s, 4, 14.0)?;

    println!("[Sheet 1] Filtered Products — 10 rows, filter Price > 50, sort Price desc");

    // ================================================================
    // Sheet 2 — AutoFilter with string condition + OR join
    // ================================================================
    let s2 = w.add_worksheet("String Filter")?;

    let headers2 = ["Name", "Department", "Level"];
    for (c, h) in headers2.iter().enumerate() {
        w.write_string(s2, 0, c as u16, h)?;
    }

    let employees: &[(&str, &str, &str)] = &[
        ("Alice", "Engineering", "Senior"),
        ("Bob", "Marketing", "Junior"),
        ("Carol", "Engineering", "Lead"),
        ("Dave", "Sales", "Senior"),
        ("Eve", "Engineering", "Junior"),
        ("Frank", "Marketing", "Senior"),
        ("Grace", "Sales", "Lead"),
        ("Hank", "Engineering", "Senior"),
    ];

    for (i, (name, dept, level)) in employees.iter().enumerate() {
        let row = 1 + i as u32;
        w.write_string(s2, row, 0, name)?;
        w.write_string(s2, row, 1, dept)?;
        w.write_string(s2, row, 2, level)?;
    }

    let last_row2 = employees.len() as u32;
    w.set_auto_filter(s2, 0, last_row2, 0, 2)?;

    // Filter column 1 (Department): "Engineering" OR "Sales".
    // Operator 0x02 = "equals" in BIFF8 DOPER encoding.
    w.add_filter_condition(
        s2,
        1,
        true, // join_or = true → show rows matching either condition
        AutoFilterConditionWrite::String {
            operator: 0x02,
            value: "Engineering".to_string(),
        },
        AutoFilterConditionWrite::String {
            operator: 0x02,
            value: "Sales".to_string(),
        },
    )?;

    // Sort by Level (column 2) ascending, then by Name (column 0) ascending.
    w.set_sort(s2, false, false, &[(2, false), (0, false)])?;

    w.set_column_width(s2, 0, 12.0)?;
    w.set_column_width(s2, 1, 16.0)?;
    w.set_column_width(s2, 2, 10.0)?;

    println!(
        "[Sheet 2] String Filter — 8 rows, filter Dept = Engineering OR Sales, sort by Level+Name"
    );

    // ================================================================
    // Sheet 3 — Multi-column conditions + boolean filter
    // ================================================================
    let s3 = w.add_worksheet("Multi-Column")?;

    let headers3 = ["Task", "Priority", "Hours", "Complete"];
    for (c, h) in headers3.iter().enumerate() {
        w.write_string(s3, 0, c as u16, h)?;
    }

    let tasks: &[(&str, &str, f64, bool)] = &[
        ("Design UI", "High", 40.0, true),
        ("Write tests", "Medium", 16.0, false),
        ("Deploy v2", "High", 8.0, false),
        ("Update docs", "Low", 4.0, true),
        ("Code review", "Medium", 12.0, true),
        ("Fix bugs", "High", 24.0, false),
    ];

    for (i, (task, prio, hours, done)) in tasks.iter().enumerate() {
        let row = 1 + i as u32;
        w.write_string(s3, row, 0, task)?;
        w.write_string(s3, row, 1, prio)?;
        w.write_number(s3, row, 2, *hours)?;
        w.write_boolean(s3, row, 3, *done)?;
    }

    let last_row3 = tasks.len() as u32;
    w.set_auto_filter(s3, 0, last_row3, 0, 3)?;

    // Filter column 2 (Hours): between 10 and 30 (>= 10 AND <= 30).
    // Operator 0x06 = ">", 0x04 = "<", 0x05 = "<=", 0x07 = ">="
    w.add_filter_condition(
        s3,
        2,
        false, // AND
        AutoFilterConditionWrite::Number {
            operator: 0x07, // >=
            value: 10.0,
        },
        AutoFilterConditionWrite::Number {
            operator: 0x05, // <=
            value: 30.0,
        },
    )?;

    // Filter column 3 (Complete): show only FALSE rows.
    w.add_filter_condition(
        s3,
        3,
        false,
        AutoFilterConditionWrite::Bool {
            operator: 0x02, // equals
            value: false,
        },
        AutoFilterConditionWrite::None,
    )?;

    // Sort by Hours descending.
    w.set_sort(s3, false, false, &[(2, true)])?;

    w.set_column_width(s3, 0, 16.0)?;
    w.set_column_width(s3, 1, 10.0)?;
    w.set_column_width(s3, 2, 10.0)?;
    w.set_column_width(s3, 3, 10.0)?;

    println!("[Sheet 3] Multi-Column — range filter on Hours + boolean filter on Complete");

    // ================================================================
    // Save
    // ================================================================
    w.save(&output_path)?;
    println!("\nSaved to: {}", output_path.display());
    println!("Open in Excel to verify filter dropdowns, conditions, and sort order.");

    Ok(())
}
