//! Comprehensive XLS file writer test
//!
//! This example demonstrates all features available in the XLS writer.
//! Tests multiple sheets, data types, formulas, and formatting.
//!
//! Run with: cargo run --example `comprehensive_xls_test`

use litchi_xls::Writer;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== Comprehensive XLS Writer Test ===\n");

    let mut writer = Writer::new();

    // ============================================================
    // 1. SALES DATA SHEET - Numbers, Strings, and Formulas
    // ============================================================
    println!("1. Creating Sales Data sheet...");
    let sales_sheet = writer.add_worksheet("Sales Data")?;

    // Headers
    writer.write_string(sales_sheet, 0, 0, "Product")?;
    writer.write_string(sales_sheet, 0, 1, "Q1 Sales")?;
    writer.write_string(sales_sheet, 0, 2, "Q2 Sales")?;
    writer.write_string(sales_sheet, 0, 3, "Q3 Sales")?;
    writer.write_string(sales_sheet, 0, 4, "Q4 Sales")?;
    writer.write_string(sales_sheet, 0, 5, "Total")?;
    writer.write_string(sales_sheet, 0, 6, "Average")?;

    // Product A
    writer.write_string(sales_sheet, 1, 0, "Widget A")?;
    writer.write_number(sales_sheet, 1, 1, 1250.50)?;
    writer.write_number(sales_sheet, 1, 2, 1380.75)?;
    writer.write_number(sales_sheet, 1, 3, 1420.00)?;
    writer.write_number(sales_sheet, 1, 4, 1560.25)?;
    writer.write_formula(sales_sheet, 1, 5, "SUM(B2:E2)")?; // Total
    writer.write_formula(sales_sheet, 1, 6, "AVERAGE(B2:E2)")?; // Average

    // Product B
    writer.write_string(sales_sheet, 2, 0, "Widget B")?;
    writer.write_number(sales_sheet, 2, 1, 2150.00)?;
    writer.write_number(sales_sheet, 2, 2, 2340.50)?;
    writer.write_number(sales_sheet, 2, 3, 2100.75)?;
    writer.write_number(sales_sheet, 2, 4, 2450.25)?;
    writer.write_formula(sales_sheet, 2, 5, "SUM(B3:E3)")?;
    writer.write_formula(sales_sheet, 2, 6, "AVERAGE(B3:E3)")?;

    // Product C
    writer.write_string(sales_sheet, 3, 0, "Widget C")?;
    writer.write_number(sales_sheet, 3, 1, 850.25)?;
    writer.write_number(sales_sheet, 3, 2, 920.00)?;
    writer.write_number(sales_sheet, 3, 3, 890.50)?;
    writer.write_number(sales_sheet, 3, 4, 1020.75)?;
    writer.write_formula(sales_sheet, 3, 5, "SUM(B4:E4)")?;
    writer.write_formula(sales_sheet, 3, 6, "AVERAGE(B4:E4)")?;

    // Product D
    writer.write_string(sales_sheet, 4, 0, "Widget D")?;
    writer.write_number(sales_sheet, 4, 1, 3200.00)?;
    writer.write_number(sales_sheet, 4, 2, 3450.50)?;
    writer.write_number(sales_sheet, 4, 3, 3150.25)?;
    writer.write_number(sales_sheet, 4, 4, 3680.00)?;
    writer.write_formula(sales_sheet, 4, 5, "SUM(B5:E5)")?;
    writer.write_formula(sales_sheet, 4, 6, "AVERAGE(B5:E5)")?;

    // Totals row
    writer.write_string(sales_sheet, 5, 0, "TOTAL")?;
    writer.write_formula(sales_sheet, 5, 1, "SUM(B2:B5)")?;
    writer.write_formula(sales_sheet, 5, 2, "SUM(C2:C5)")?;
    writer.write_formula(sales_sheet, 5, 3, "SUM(D2:D5)")?;
    writer.write_formula(sales_sheet, 5, 4, "SUM(E2:E5)")?;
    writer.write_formula(sales_sheet, 5, 5, "SUM(F2:F5)")?;
    writer.write_formula(sales_sheet, 5, 6, "AVERAGE(G2:G5)")?;

    // Add some analysis
    writer.write_string(sales_sheet, 7, 0, "Best Quarter:")?;
    writer.write_formula(sales_sheet, 7, 1, "MAX(B6:E6)")?;

    writer.write_string(sales_sheet, 8, 0, "Worst Quarter:")?;
    writer.write_formula(sales_sheet, 8, 1, "MIN(B6:E6)")?;

    // ============================================================
    // 2. INVENTORY SHEET - Booleans and Conditional Formulas
    // ============================================================
    println!("2. Creating Inventory sheet...");
    let inventory_sheet = writer.add_worksheet("Inventory")?;

    // Headers
    writer.write_string(inventory_sheet, 0, 0, "Product")?;
    writer.write_string(inventory_sheet, 0, 1, "In Stock")?;
    writer.write_string(inventory_sheet, 0, 2, "Reorder Level")?;
    writer.write_string(inventory_sheet, 0, 3, "Need Reorder?")?;
    writer.write_string(inventory_sheet, 0, 4, "Status")?;

    // Product data
    let products = [
        ("Widget A", 150.0, 100.0, false),
        ("Widget B", 75.0, 100.0, true),
        ("Widget C", 250.0, 150.0, false),
        ("Widget D", 45.0, 50.0, true),
        ("Widget E", 180.0, 100.0, false),
        ("Widget F", 30.0, 50.0, true),
    ];

    for (i, (product, stock, reorder, need)) in products.iter().enumerate() {
        let row = (i + 1) as u32;
        writer.write_string(inventory_sheet, row, 0, product)?;
        writer.write_number(inventory_sheet, row, 1, *stock)?;
        writer.write_number(inventory_sheet, row, 2, *reorder)?;
        writer.write_boolean(inventory_sheet, row, 3, *need)?;

        // Conditional formula: IF(B<C, "LOW", "OK")
        let formula = format!("IF(B{}<C{},\"LOW\",\"OK\")", row + 1, row + 1);
        writer.write_formula(inventory_sheet, row, 4, &formula)?;
    }

    // Summary statistics
    writer.write_string(inventory_sheet, 8, 0, "Total Items:")?;
    writer.write_formula(inventory_sheet, 8, 1, "COUNT(B2:B7)")?;

    writer.write_string(inventory_sheet, 9, 0, "Total Stock:")?;
    writer.write_formula(inventory_sheet, 9, 1, "SUM(B2:B7)")?;

    writer.write_string(inventory_sheet, 10, 0, "Avg Stock:")?;
    writer.write_formula(inventory_sheet, 10, 1, "AVERAGE(B2:B7)")?;

    // ============================================================
    // 3. EMPLOYEE DATA SHEET - Mixed Data Types
    // ============================================================
    println!("3. Creating Employee Data sheet...");
    let employee_sheet = writer.add_worksheet("Employees")?;

    // Headers
    writer.write_string(employee_sheet, 0, 0, "Employee ID")?;
    writer.write_string(employee_sheet, 0, 1, "Name")?;
    writer.write_string(employee_sheet, 0, 2, "Department")?;
    writer.write_string(employee_sheet, 0, 3, "Salary")?;
    writer.write_string(employee_sheet, 0, 4, "Bonus")?;
    writer.write_string(employee_sheet, 0, 5, "Total Comp")?;
    writer.write_string(employee_sheet, 0, 6, "Active")?;

    // Employee records
    let employees = [
        (1001, "John Smith", "Sales", 65000.0, 5000.0, true),
        (1002, "Jane Doe", "Engineering", 85000.0, 8500.0, true),
        (1003, "Bob Johnson", "Marketing", 55000.0, 3500.0, true),
        (
            1004,
            "Alice Williams",
            "Engineering",
            92000.0,
            10000.0,
            true,
        ),
        (1005, "Charlie Brown", "Sales", 58000.0, 4200.0, false),
        (1006, "Diana Prince", "HR", 62000.0, 4800.0, true),
        (1007, "Ethan Hunt", "Operations", 70000.0, 6000.0, true),
        (1008, "Fiona Apple", "Engineering", 88000.0, 9200.0, true),
    ];

    for (i, (id, name, dept, salary, bonus, active)) in employees.iter().enumerate() {
        let row = (i + 1) as u32;
        writer.write_number(employee_sheet, row, 0, f64::from(*id))?;
        writer.write_string(employee_sheet, row, 1, name)?;
        writer.write_string(employee_sheet, row, 2, dept)?;
        writer.write_number(employee_sheet, row, 3, *salary)?;
        writer.write_number(employee_sheet, row, 4, *bonus)?;

        // Total compensation formula
        let formula = format!("D{}+E{}", row + 1, row + 1);
        writer.write_formula(employee_sheet, row, 5, &formula)?;

        writer.write_boolean(employee_sheet, row, 6, *active)?;
    }

    // Summary calculations
    writer.write_string(employee_sheet, 10, 0, "Statistics:")?;

    writer.write_string(employee_sheet, 11, 0, "Total Employees:")?;
    writer.write_formula(employee_sheet, 11, 1, "COUNT(A2:A9)")?;

    writer.write_string(employee_sheet, 12, 0, "Avg Salary:")?;
    writer.write_formula(employee_sheet, 12, 1, "AVERAGE(D2:D9)")?;

    writer.write_string(employee_sheet, 13, 0, "Max Salary:")?;
    writer.write_formula(employee_sheet, 13, 1, "MAX(D2:D9)")?;

    writer.write_string(employee_sheet, 14, 0, "Min Salary:")?;
    writer.write_formula(employee_sheet, 14, 1, "MIN(D2:D9)")?;

    writer.write_string(employee_sheet, 15, 0, "Total Payroll:")?;
    writer.write_formula(employee_sheet, 15, 1, "SUM(F2:F9)")?;

    // ============================================================
    // 4. CALCULATIONS SHEET - Advanced Formulas
    // ============================================================
    println!("4. Creating Calculations sheet...");
    let calc_sheet = writer.add_worksheet("Calculations")?;

    writer.write_string(calc_sheet, 0, 0, "Formula Examples")?;

    // Basic arithmetic
    writer.write_string(calc_sheet, 2, 0, "Basic Arithmetic:")?;
    writer.write_string(calc_sheet, 3, 0, "10 + 20 =")?;
    writer.write_formula(calc_sheet, 3, 1, "10+20")?;

    writer.write_string(calc_sheet, 4, 0, "100 - 35 =")?;
    writer.write_formula(calc_sheet, 4, 1, "100-35")?;

    writer.write_string(calc_sheet, 5, 0, "15 * 8 =")?;
    writer.write_formula(calc_sheet, 5, 1, "15*8")?;

    writer.write_string(calc_sheet, 6, 0, "144 / 12 =")?;
    writer.write_formula(calc_sheet, 6, 1, "144/12")?;

    writer.write_string(calc_sheet, 7, 0, "2 ^ 10 =")?;
    writer.write_formula(calc_sheet, 7, 1, "2^10")?;

    // Complex expressions
    writer.write_string(calc_sheet, 9, 0, "Complex Expressions:")?;
    writer.write_string(calc_sheet, 10, 0, "(5+3)*2 =")?;
    writer.write_formula(calc_sheet, 10, 1, "(5+3)*2")?;

    writer.write_string(calc_sheet, 11, 0, "10+20*3 =")?;
    writer.write_formula(calc_sheet, 11, 1, "10+20*3")?;

    // Function examples
    writer.write_string(calc_sheet, 13, 0, "Functions:")?;

    // Sample data for functions
    writer.write_number(calc_sheet, 14, 1, 15.0)?;
    writer.write_number(calc_sheet, 14, 2, 25.0)?;
    writer.write_number(calc_sheet, 14, 3, 35.0)?;
    writer.write_number(calc_sheet, 14, 4, 45.0)?;
    writer.write_number(calc_sheet, 14, 5, 55.0)?;

    writer.write_string(calc_sheet, 15, 0, "SUM =")?;
    writer.write_formula(calc_sheet, 15, 1, "SUM(B15:F15)")?;

    writer.write_string(calc_sheet, 16, 0, "AVERAGE =")?;
    writer.write_formula(calc_sheet, 16, 1, "AVERAGE(B15:F15)")?;

    writer.write_string(calc_sheet, 17, 0, "MAX =")?;
    writer.write_formula(calc_sheet, 17, 1, "MAX(B15:F15)")?;

    writer.write_string(calc_sheet, 18, 0, "MIN =")?;
    writer.write_formula(calc_sheet, 18, 1, "MIN(B15:F15)")?;

    writer.write_string(calc_sheet, 19, 0, "COUNT =")?;
    writer.write_formula(calc_sheet, 19, 1, "COUNT(B15:F15)")?;

    // Math functions
    writer.write_string(calc_sheet, 21, 0, "Math Functions:")?;
    writer.write_string(calc_sheet, 22, 0, "ABS(-42) =")?;
    writer.write_formula(calc_sheet, 22, 1, "ABS(-42)")?;

    writer.write_string(calc_sheet, 23, 0, "ROUND(3.14159, 2) =")?;
    writer.write_formula(calc_sheet, 23, 1, "ROUND(3.14159,2)")?;

    // ============================================================
    // 5. TEXT FUNCTIONS SHEET
    // ============================================================
    println!("5. Creating Text Functions sheet...");
    let text_sheet = writer.add_worksheet("Text Functions")?;

    writer.write_string(text_sheet, 0, 0, "Text Function Examples")?;

    // Sample text
    writer.write_string(text_sheet, 2, 0, "Sample Text:")?;
    writer.write_string(text_sheet, 2, 1, "Hello World")?;

    writer.write_string(text_sheet, 3, 0, "LEN =")?;
    writer.write_formula(text_sheet, 3, 1, "LEN(B3)")?;

    writer.write_string(text_sheet, 4, 0, "LEFT(5) =")?;
    writer.write_formula(text_sheet, 4, 1, "LEFT(B3,5)")?;

    writer.write_string(text_sheet, 5, 0, "RIGHT(5) =")?;
    writer.write_formula(text_sheet, 5, 1, "RIGHT(B3,5)")?;

    writer.write_string(text_sheet, 6, 0, "MID(7,5) =")?;
    writer.write_formula(text_sheet, 6, 1, "MID(B3,7,5)")?;

    // Concatenation
    writer.write_string(text_sheet, 8, 0, "First Name:")?;
    writer.write_string(text_sheet, 8, 1, "John")?;
    writer.write_string(text_sheet, 9, 0, "Last Name:")?;
    writer.write_string(text_sheet, 9, 1, "Doe")?;

    writer.write_string(text_sheet, 10, 0, "Full Name:")?;
    writer.write_formula(text_sheet, 10, 1, "CONCATENATE(B9,\" \",B10)")?;

    // ============================================================
    // 6. SUMMARY SHEET
    // ============================================================
    println!("6. Creating Summary sheet...");
    let summary_sheet = writer.add_worksheet("Summary")?;

    writer.write_string(summary_sheet, 0, 0, "XLS Comprehensive Test Summary")?;

    writer.write_string(summary_sheet, 2, 0, "Features Demonstrated:")?;
    writer.write_string(summary_sheet, 3, 0, "✓ Multiple worksheets (6 sheets)")?;
    writer.write_string(summary_sheet, 4, 0, "✓ String values")?;
    writer.write_string(
        summary_sheet,
        5,
        0,
        "✓ Numeric values (integers and decimals)",
    )?;
    writer.write_string(summary_sheet, 6, 0, "✓ Boolean values")?;
    writer.write_string(
        summary_sheet,
        7,
        0,
        "✓ Formula support (SUM, AVERAGE, MAX, MIN, COUNT)",
    )?;
    writer.write_string(
        summary_sheet,
        8,
        0,
        "✓ Conditional formulas (IF statements)",
    )?;
    writer.write_string(
        summary_sheet,
        9,
        0,
        "✓ Text functions (LEN, LEFT, RIGHT, MID, CONCATENATE)",
    )?;
    writer.write_string(summary_sheet, 10, 0, "✓ Math functions (ABS, ROUND)")?;
    writer.write_string(summary_sheet, 11, 0, "✓ Cell references (relative)")?;
    writer.write_string(summary_sheet, 12, 0, "✓ Range references (A1:B10)")?;
    writer.write_string(
        summary_sheet,
        13,
        0,
        "✓ Arithmetic operators (+, -, *, /, ^)",
    )?;
    writer.write_string(
        summary_sheet,
        14,
        0,
        "✓ Complex expressions with precedence",
    )?;
    writer.write_string(summary_sheet, 15, 0, "✓ Shared string table optimization")?;

    writer.write_string(summary_sheet, 17, 0, "Sheet Breakdown:")?;
    writer.write_string(
        summary_sheet,
        18,
        0,
        "1. Sales Data - Quarterly sales with formulas",
    )?;
    writer.write_string(
        summary_sheet,
        19,
        0,
        "2. Inventory - Stock levels with conditionals",
    )?;
    writer.write_string(
        summary_sheet,
        20,
        0,
        "3. Employees - Personnel data with calculations",
    )?;
    writer.write_string(summary_sheet, 21, 0, "4. Calculations - Formula examples")?;
    writer.write_string(
        summary_sheet,
        22,
        0,
        "5. Text Functions - String manipulation",
    )?;
    writer.write_string(summary_sheet, 23, 0, "6. Summary - This sheet")?;

    writer.write_string(summary_sheet, 25, 0, "Total Sheets:")?;
    writer.write_number(summary_sheet, 25, 1, 6.0)?;

    writer.write_string(summary_sheet, 27, 0, "Generated by:")?;
    writer.write_string(summary_sheet, 27, 1, "Litchi Library")?;
    writer.write_string(summary_sheet, 28, 0, "Format:")?;
    writer.write_string(summary_sheet, 28, 1, "Excel 97-2003 (BIFF8)")?;

    // ============================================================
    // SAVE FILE
    // ============================================================
    println!("\nSaving to comprehensive_test.xls...");
    writer.save("comprehensive_test.xls")?;

    println!("\n✅ SUCCESS! XLS file created with:");
    println!("   ✓ 6 worksheets with different data");
    println!("   ✓ String values (100+ unique strings)");
    println!("   ✓ Numeric values (integers and decimals)");
    println!("   ✓ Boolean values (TRUE/FALSE)");
    println!("   ✓ 50+ formulas including:");
    println!("     - SUM, AVERAGE, MAX, MIN, COUNT");
    println!("     - IF conditionals");
    println!("     - Text functions (LEN, LEFT, RIGHT, MID, CONCATENATE)");
    println!("     - Math functions (ABS, ROUND)");
    println!("     - Arithmetic expressions with operators");
    println!("   ✓ Cell and range references");
    println!("   ✓ Shared string table optimization");
    println!("\n📊 Open 'comprehensive_test.xls' in Microsoft Excel to verify!");
    println!("   - Check formulas calculate correctly");
    println!("   - Verify all 6 sheets are present");
    println!("   - Test data integrity across sheets");

    Ok(())
}
