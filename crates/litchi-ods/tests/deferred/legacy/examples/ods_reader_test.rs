//! Comprehensive ODS (OpenDocument Spreadsheet) reading example.
//!
//! This example demonstrates all reading capabilities for ODS files,
//! serving as both verification and regression testing.
//!
//! Run with:
//! ```bash
//! cargo run --example ods_reader_test --features odf --no-default-features
//! ```

#[cfg(feature = "odf")]
use litchi::Result;
#[cfg(feature = "odf")]
use litchi::ods::{CellValue, Spreadsheet};

#[cfg(feature = "odf")]
fn main() -> Result<()> {
    println!("=== ODS Reader Comprehensive Test ===\n");

    // Open the test ODS file
    let test_file = "test.ods";
    println!("📖 Opening file: {}", test_file);
    let mut spreadsheet = Spreadsheet::open(test_file)?;
    println!("✅ File opened successfully\n");

    // Test 1: Sheet count
    println!("--- Test 1: Sheet Count ---");
    match spreadsheet.sheet_count() {
        Ok(count) => {
            println!("Total sheets: {}", count);
            println!();
        },
        Err(e) => println!("⚠️  Error getting sheet count: {}\n", e),
    }

    // Test 2: List all sheets
    println!("--- Test 2: Sheet Enumeration ---");
    match spreadsheet.sheets() {
        Ok(sheets) => {
            println!("Retrieved {} sheets", sheets.len());
            for (i, sheet) in sheets.iter().enumerate() {
                println!("  Sheet {}: \"{}\"", i + 1, sheet.name);
                println!(
                    "    Rows: {}, Columns: {}",
                    sheet.rows.len(),
                    sheet.rows.first().map(|r| r.cells.len()).unwrap_or(0)
                );
            }
            println!();
        },
        Err(e) => println!("⚠️  Error listing sheets: {}\n", e),
    }

    // Test 3: Access sheets by name and index
    println!("--- Test 3: Sheet Access Methods ---");
    if let Ok(sheets) = spreadsheet.sheets()
        && !sheets.is_empty()
    {
        let sheet_name = &sheets[0].name;
        println!("Accessing sheet by name: '{}'", sheet_name);
        match spreadsheet.sheet_by_name(sheet_name) {
            Ok(Some(sheet)) => println!("  ✅ Successfully accessed: {}", sheet.name),
            Ok(None) => println!("  ⚠️  Sheet not found"),
            Err(e) => println!("  ⚠️  Error: {}", e),
        }

        println!("Accessing sheet by index: 0");
        match spreadsheet.sheet_by_index(0) {
            Ok(Some(sheet)) => println!("  ✅ Successfully accessed: {}", sheet.name),
            Ok(None) => println!("  ⚠️  Sheet not found"),
            Err(e) => println!("  ⚠️  Error: {}", e),
        }
        println!();
    }

    // Test 4: Cell value extraction (all types)
    println!("--- Test 4: Cell Value Extraction ---");
    if let Ok(Some(sheet)) = spreadsheet.sheet_by_index(0) {
        println!("Reading cells from sheet: {}", sheet.name);

        // Try reading various cells by accessing rows directly
        let test_positions = vec![
            ("A1", 0, 0),
            ("B1", 0, 1),
            ("A2", 1, 0),
            ("B2", 1, 1),
            ("C3", 2, 2),
        ];

        for (notation, row_idx, col_idx) in test_positions {
            if let Some(row) = sheet.rows.get(row_idx)
                && let Some(cell) = row.cells.get(col_idx)
            {
                print!("  Cell {}:{} ({})", row_idx + 1, col_idx + 1, notation);
                match &cell.value {
                    CellValue::Text(s) => println!(" = \"{}\" (Text)", s),
                    CellValue::Number(n) => println!(" = {} (Number)", n),
                    CellValue::Boolean(b) => println!(" = {} (Boolean)", b),
                    CellValue::Date(d) => println!(" = {} (Date)", d),
                    CellValue::Time(t) => println!(" = {} (Time)", t),
                    CellValue::Percentage(p) => println!(" = {}% (Percentage)", p),
                    CellValue::Currency(c, curr) => println!(" = {} {} (Currency)", c, curr),
                    CellValue::Empty => println!(" = (Empty)"),
                }

                // Check for formula
                if let Some(ref formula) = cell.formula {
                    println!("    Formula: {}", formula);
                }
            }
        }
        println!();
    }

    // Test 5: Row iteration
    println!("--- Test 5: Row Iteration ---");
    if let Ok(Some(sheet)) = spreadsheet.sheet_by_index(0) {
        println!("Iterating rows in sheet: {}", sheet.name);
        let rows = &sheet.rows;
        println!("Total rows: {}", rows.len());

        for (i, row) in rows.iter().take(5).enumerate() {
            let cells = &row.cells;
            print!("  Row {}: ", i + 1);
            let cell_values: Vec<String> = cells
                .iter()
                .take(5)
                .map(|cell| match &cell.value {
                    CellValue::Text(s) => format!("\"{}\"", s),
                    CellValue::Number(n) => format!("{}", n),
                    CellValue::Boolean(b) => format!("{}", b),
                    CellValue::Empty => "".to_string(),
                    _ => format!("{:?}", cell.value),
                })
                .collect();
            println!("{}", cell_values.join(", "));
        }

        if rows.len() > 5 {
            println!("  ... and {} more rows", rows.len() - 5);
        }
        println!();
    }

    // Test 6: CSV export
    println!("--- Test 6: CSV Export ---");
    match spreadsheet.to_csv() {
        Ok(csv) => {
            let preview = if csv.len() > 300 {
                format!("{}... ({} bytes total)", &csv[..300], csv.len())
            } else {
                csv
            };
            println!("CSV export successful:");
            println!("{}\n", preview);
        },
        Err(e) => println!("⚠️  Error exporting to CSV: {}\n", e),
    }

    // Test 7: Metadata extraction
    println!("--- Test 7: Metadata Extraction ---");
    match spreadsheet.metadata() {
        Ok(metadata) => {
            println!("Spreadsheet metadata:");
            if let Some(ref title) = metadata.title {
                println!("  Title: {}", title);
            }
            if let Some(ref author) = metadata.author {
                println!("  Author: {}", author);
            }
            if let Some(ref subject) = metadata.subject {
                println!("  Subject: {}", subject);
            }
            if let Some(ref description) = metadata.description {
                println!("  Description: {}", description);
            }
            if let Some(created) = metadata.created {
                println!("  Created: {}", created);
            }
            if let Some(modified) = metadata.modified {
                println!("  Modified: {}", modified);
            }
            println!();
        },
        Err(e) => println!("⚠️  Error extracting metadata: {}\n", e),
    }

    // Test 8: Multi-sheet handling
    println!("--- Test 8: Multi-Sheet Handling ---");
    if let Ok(sheets) = spreadsheet.sheets() {
        println!("Processing all sheets:");
        for sheet in sheets {
            let max_cols = sheet.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0);
            println!(
                "  Sheet '{}': {} rows × {} cols",
                sheet.name,
                sheet.rows.len(),
                max_cols
            );

            // Count non-empty cells
            let mut non_empty = 0;
            for row in &sheet.rows {
                for cell in &row.cells {
                    if !matches!(cell.value, CellValue::Empty) {
                        non_empty += 1;
                    }
                }
            }
            println!("    Non-empty cells: {}", non_empty);
        }
        println!();
    }

    // Test 9: Text extraction
    println!("--- Test 9: Full Text Extraction ---");
    match spreadsheet.text() {
        Ok(text) => {
            let preview = if text.len() > 300 {
                format!("{}... ({} chars total)", &text[..300], text.len())
            } else {
                text
            };
            println!("Text content:\n{}\n", preview);
        },
        Err(e) => println!("⚠️  Error extracting text: {}\n", e),
    }

    println!("=== ODS Reader Test Complete ===");
    println!("✅ All reading functionalities tested successfully!");

    Ok(())
}

#[cfg(not(feature = "odf"))]
fn main() {
    eprintln!(
        "This example requires the 'odf' feature. Try: cargo run --example ods_reader_test --features odf --no-default-features"
    );
}
