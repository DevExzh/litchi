//! API demonstration for text format functions
//!
//! This example shows how to use the public API functions for opening
//! text format workbooks through the sheet module.

use litchi::sheet::CellValue;
use litchi::sheet::functions::*;

type ExampleResult<T> = Result<T, Box<dyn std::error::Error + Send + Sync>>;

fn main() -> ExampleResult<()> {
    println!("=== Text Formats API Demo ===\n");

    // Create sample CSV data
    create_sample_csv()?;

    // Test each format through the public API
    test_csv_api()?;
    test_tsv_api();
    test_prn_api();
    test_sylk_api();
    test_dif_api();
    test_fixed_width_api();

    println!("\n=== API demo completed! ===");
    Ok(())
}

fn create_sample_csv() -> ExampleResult<()> {
    println!("Creating sample CSV file...");

    let csv_data = "Product,Price,Quantity,Total\nLaptop,999.99,2,1999.98\nMouse,29.99,5,149.95\nKeyboard,79.99,3,239.97\n";

    std::fs::write("sample.csv", csv_data)?;
    println!("  ✓ Created sample.csv");

    // Create sample TSV
    let tsv_data =
        "Product\tPrice\tQuantity\tTotal\nLaptop\t999.99\t2\t1999.98\nMouse\t29.99\t5\t149.95\n";
    std::fs::write("sample.tsv", tsv_data)?;
    println!("  ✓ Created sample.tsv");

    // Create sample SYLK
    let sylk_data = "ID;PWXL;N;E\nB;Y4;X4\nC;Y1;X1;K\"Product\"\nC;Y1;X2;K\"Price\"\nC;Y1;X3;K\"Quantity\"\nC;Y1;X4;K\"Total\"\nC;Y2;X1;K\"Laptop\"\nC;Y2;X2;K999.99\nC;Y2;X3;K2\nC;Y2;X4;E999.99*2\nE\n";
    std::fs::write("sample.slk", sylk_data)?;
    println!("  ✓ Created sample.slk");

    // Create sample DIF
    let dif_data = "TABLE\n0,1\n\"DEMO\"\nVECTORS\n0,4\n\"\"\nTUPLES\n0,2\n\"\"\nDATA\n0,0\n\"\"\n-1,0\nBOT\n1,0\n\"Product\"\n1,0\n\"Price\"\n1,0\n\"Quantity\"\n1,0\n\"Total\"\n-1,0\nBOT\n1,0\n\"Laptop\"\n0,999.99\nV\n0,2\nV\n0,1999.98\nV\n-1,0\nEOD\n";
    std::fs::write("sample.dif", dif_data)?;
    println!("  ✓ Created sample.dif");

    // Create sample fixed-width PRN
    let prn_data = "Product    Price    Quantity  Total    \nLaptop     999.99   2         1999.98  \nMouse      29.99    5         149.95   \n";
    std::fs::write("sample.prn", prn_data)?;
    println!("  ✓ Created sample.prn");

    Ok(())
}

fn test_csv_api() -> ExampleResult<()> {
    println!("\nTesting CSV API...");

    // Open CSV workbook
    let workbook = open_csv_workbook("sample.csv")?;

    // Get worksheet
    let worksheet = workbook.worksheet_by_index(0)?;

    println!("  ✓ Opened CSV workbook");
    println!("  ✓ Worksheet name: {}", worksheet.name());

    // Read data
    let mut row_count = 0;
    let mut col_count = 0;

    for row_idx in 0..worksheet.row_count() {
        let row = worksheet.row(row_idx)?;
        if !row.is_empty() {
            row_count += 1;
            col_count = col_count.max(row.len());

            if row_count <= 2 {
                // Print first 2 rows
                let row_data: Vec<String> = row.iter().map(|cell| format!("{:?}", cell)).collect();
                println!("    Row {}: {:?}", row_idx, row_data);
            }
        }
    }

    println!("  ✓ Read {} rows, {} columns", row_count, col_count);
    Ok(())
}

fn test_tsv_api() {
    println!("\nTesting TSV API...");

    match open_tsv_workbook("sample.tsv") {
        Ok(workbook) => {
            let worksheet = workbook.worksheet_by_index(0).unwrap();
            println!("  ✓ Opened TSV workbook: {}", worksheet.name());
            println!("  ✓ Rows: {}", worksheet.row_count());
        },
        Err(e) => println!("  ❌ Failed to open TSV: {}", e),
    }
}

fn test_prn_api() {
    println!("\nTesting PRN API...");

    // Test delimited PRN
    match open_prn_workbook("sample.csv") {
        // Using CSV as PRN for demo
        Ok(workbook) => {
            let worksheet = workbook.worksheet_by_index(0).unwrap();
            println!("  ✓ Opened PRN (delimited) workbook");
            println!("  ✓ Rows: {}", worksheet.row_count());
        },
        Err(e) => println!("  ❌ Failed to open PRN: {}", e),
    }

    // Test fixed-width PRN
    match open_fixed_width_workbook("sample.prn") {
        Ok(workbook) => {
            let worksheet = workbook.worksheet_by_index(0).unwrap();
            println!("  ✓ Opened PRN (fixed-width) workbook");
            println!("  ✓ Rows: {}", worksheet.row_count());
        },
        Err(e) => println!("  ❌ Failed to open fixed-width PRN: {}", e),
    }
}

fn test_sylk_api() {
    println!("\nTesting SYLK API...");

    match open_sylk_workbook("sample.slk") {
        Ok(workbook) => {
            let worksheet = workbook.worksheet_by_index(0).unwrap();
            println!("  ✓ Opened SYLK workbook");
            println!("  ✓ Rows: {}", worksheet.row_count());

            // Check for formulas
            for row_idx in 0..worksheet.row_count() {
                let row = worksheet.row(row_idx).unwrap();
                for (col_idx, cell) in row.iter().enumerate() {
                    if let CellValue::Formula { formula, .. } = cell {
                        println!("    Formula at ({}, {}): {}", row_idx, col_idx, formula);
                    }
                }
            }
        },
        Err(e) => println!("  ❌ Failed to open SYLK: {}", e),
    }
}

fn test_dif_api() {
    println!("\nTesting DIF API...");

    match open_dif_workbook("sample.dif") {
        Ok(workbook) => {
            let worksheet = workbook.worksheet_by_index(0).unwrap();
            println!("  ✓ Opened DIF workbook");
            println!("  ✓ Rows: {}", worksheet.row_count());
        },
        Err(e) => println!("  ❌ Failed to open DIF: {}", e),
    }
}

fn test_fixed_width_api() {
    println!("\nTesting Fixed-width API...");

    match open_fixed_width_workbook("sample.prn") {
        Ok(workbook) => {
            let worksheet = workbook.worksheet_by_index(0).unwrap();
            println!("  ✓ Opened fixed-width workbook");
            println!("  ✓ Rows: {}", worksheet.row_count());

            // Show first row to verify fixed-width parsing
            if let Ok(row) = worksheet.row(0) {
                let row_data: Vec<String> = row.iter().map(|cell| format!("{:?}", cell)).collect();
                println!("    First row: {:?}", row_data);
            }
        },
        Err(e) => println!("  ❌ Failed to open fixed-width: {}", e),
    }
}
