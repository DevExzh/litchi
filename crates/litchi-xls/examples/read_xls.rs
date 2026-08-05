//! Read a legacy Excel `.xls` file and print sheet/cell information.
//!
//! Run with:
//!     cargo run -p litchi-xls --example read_xls
//!     cargo run -p litchi-xls --example read_xls -- path/to/file.xls

use litchi_core::sheet::{CellValue, WorkbookTrait, Worksheet as _};
use litchi_xls::Workbook;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "test-data/ole/xls/Simple.xls".to_string());

    println!("Opening XLS: {}", path);

    let reader = File::open(&path)?;
    let workbook = Workbook::new(reader)?;

    println!("\n=== Workbook ===");
    println!("Worksheet count : {}", workbook.worksheet_count());
    println!("1904 date system: {}", workbook.is_1904_date_system());

    let names: Vec<String> = workbook.worksheet_names().to_vec();
    for (idx, name) in names.iter().enumerate() {
        println!("\n--- Sheet [{}]: \"{}\" ---", idx, name);

        let sheet = workbook.xls_worksheet(idx)?;
        println!(
            "Rows x Cols: {} x {}",
            sheet.row_count(),
            sheet.column_count()
        );
        match sheet.dimensions() {
            Some((min_r, min_c, max_r, max_c)) => println!(
                "Dimensions : ({}, {}) - ({}, {})",
                min_r, min_c, max_r, max_c
            ),
            None => println!("Dimensions : <empty>"),
        }

        // Print up to 10 non-empty cells as a sample.
        let mut shown = 0usize;
        let mut iter = sheet.cells();
        while let Some(cell_result) = iter.next() {
            let cell =
                cell_result.map_err(|e| -> Box<dyn std::error::Error> { e.to_string().into() })?;
            if cell.is_empty() {
                continue;
            }
            println!(
                "  {}({},{}) = {}",
                cell.coordinate(),
                cell.row(),
                cell.column(),
                format_value(cell.value())
            );
            shown += 1;
            if shown >= 10 {
                println!("  ... [showing first 10 non-empty cells only]");
                break;
            }
        }
        if shown == 0 {
            println!("  (no non-empty cells)");
        }

        // Show extras if present.
        let merged = sheet.merged_cells();
        if !merged.is_empty() {
            println!("Merged ranges: {}", merged.len());
        }
        let hyperlinks = sheet.hyperlinks();
        if !hyperlinks.is_empty() {
            println!("Hyperlinks   : {}", hyperlinks.len());
        }
        let comments = sheet.comments();
        if !comments.is_empty() {
            println!("Comments     : {}", comments.len());
        }
    }

    Ok(())
}

fn format_value(v: &CellValue) -> String {
    match v {
        CellValue::Empty => "<empty>".to_string(),
        CellValue::Bool(b) => format!("{}", b),
        CellValue::Int(i) => format!("{}", i),
        CellValue::Float(f) => format!("{}", f),
        CellValue::String(s) => format!("\"{}\"", s),
        CellValue::DateTime(d) => format!("date({})", d),
        CellValue::Error(e) => format!("#ERR({})", e),
        CellValue::Formula { .. } => "<formula>".to_string(),
    }
}
